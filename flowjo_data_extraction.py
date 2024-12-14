import sys
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                            QHBoxLayout, QLabel, QPushButton, QComboBox, 
                            QRadioButton, QButtonGroup, QListWidget, QFileDialog,
                            QGroupBox, QCheckBox, QMessageBox)
from PyQt6.QtCore import Qt
import pandas as pd
import numpy as np
from scipy import stats
import pyperclip
import traceback
import warnings

class FlowDataProcessor(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Flow Cytometry Data Processor")
        self.setMinimumWidth(800)
    
        # Data storage
        self.sample_map = None
        self.group_map = None
        self.flow_data = None
        self.processed_data = None
        self.sample_well_data = {}
        self.group_well_data = {}
        self.sample_order = None  # Store sample order from column 14
        self.group_order = None   # Store group order from column 14
    
        # Create main widget and layout
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout(main_widget)
    
        # Create sections and add them to main layout
        main_layout.addWidget(self.create_file_loading_section())
        main_layout.addWidget(self.create_data_display_section())
        main_layout.addWidget(self.create_analysis_options_section())
        main_layout.addWidget(self.create_export_section())
    
        # Add credit text at bottom
        credit_label = QLabel("Inspired by the legendary Dan Piraner")
        credit_label.setStyleSheet("color: gray; font-style: italic;")
        credit_label.setAlignment(Qt.AlignmentFlag.AlignRight)
        main_layout.addWidget(credit_label)
    
        self.update_ui_state()

    def create_file_loading_section(self):
        group_box = QGroupBox("Load Files")
        layout = QVBoxLayout()
        
        # Sample Map
        sample_layout = QHBoxLayout()
        self.load_sample_btn = QPushButton("Load Sample Map")
        self.load_sample_btn.clicked.connect(self.load_sample_map)
        self.sample_status = QLabel("No Sample Map")
        self.sample_status.setStyleSheet("color: red")
        sample_layout.addWidget(self.load_sample_btn)
        sample_layout.addWidget(self.sample_status)
        sample_layout.addStretch()
        layout.addLayout(sample_layout)
        
        # Group Map
        group_layout = QHBoxLayout()
        self.load_group_btn = QPushButton("Load Group Map")
        self.load_group_btn.clicked.connect(self.load_group_map)
        self.group_status = QLabel("No Group Map")
        self.group_status.setStyleSheet("color: red")
        group_layout.addWidget(self.load_group_btn)
        group_layout.addWidget(self.group_status)
        group_layout.addStretch()
        layout.addLayout(group_layout)
        
        # Flow Data
        flow_layout = QHBoxLayout()
        self.load_flow_btn = QPushButton("Load Flow Data")
        self.load_flow_btn.clicked.connect(self.load_flowjo_data)
        self.flow_status = QLabel("No Flow Data")
        self.flow_status.setStyleSheet("color: red")
        flow_layout.addWidget(self.load_flow_btn)
        flow_layout.addWidget(self.flow_status)
        flow_layout.addStretch()
        layout.addLayout(flow_layout)
        
        group_box.setLayout(layout)
        return group_box

    def load_sample_map(self):
        filename, _ = QFileDialog.getOpenFileName(self, "Load Sample Map",
                                                "", "Excel Files (*.xlsx *.xls)")
        if filename:
            try:
                # Read the entire Excel file
                df = pd.read_excel(filename)
                
                # Extract sample order from column 14 (index 13), starting from row 1 excluding header
                self.sample_order = df.iloc[0:, 13].dropna().tolist()
                print(f"Loaded sample order from column 14: {self.sample_order}")
                
                # Process the plate map as before
                self.sample_map, self.sample_well_data = self.process_plate_map(df)
                self.sample_status.setText("LOADED")
                self.sample_status.setStyleSheet("color: green")
                
                # Update sample options
                if self.sample_radio.isChecked():
                    self.update_filter_list()
                
                self.update_ui_state()
                
            except Exception as e:
                self.sample_status.setText("LOAD ERROR")
                self.sample_status.setStyleSheet("color: red")
                QMessageBox.critical(self, "Error", f"Error loading Sample Map: {str(e)}")

    def process_data(self):
        if not all([self.sample_map is not None, self.group_map is not None, self.flow_data is not None]):
            QMessageBox.critical(self, "Error", "Please load all required files first")
            return None

        selected_measurement = self.measurement_combo.currentText()
        if not selected_measurement:
            QMessageBox.critical(self, "Error", "Please select a measurement")
            return None

        try:
            # Merge flow data with sample information first
            merged_data = self.flow_data.merge(
                self.sample_map,
                on='Well',
                how='left'
            )
            merged_data = merged_data.rename(columns={'Value': 'Sample'})
    
            # Then merge with group information
            merged_data = merged_data.merge(
                self.group_map,
                on='Well',
                how='left'
            )
            merged_data = merged_data.rename(columns={'Value': 'Group'})

            # Get selected options from QListWidget
            selected_items = self.filter_list.selectedItems()
            if not selected_items:
                selected_options = ["All"]
            else:
                selected_options = [item.text() for item in selected_items]
    
            # Apply filters based on selection type
            if self.sample_radio.isChecked() and "All" not in selected_options:
                merged_data = merged_data[merged_data['Sample'].isin(selected_options)]
            elif self.group_radio.isChecked() and "All" not in selected_options:
                merged_data = merged_data[merged_data['Group'].isin(selected_options)]
        
            merged_data[selected_measurement] = pd.to_numeric(merged_data[selected_measurement], errors='coerce').round(2)
    
            if self.individual_radio.isChecked():
                # Get groups in the specified order
                if self.group_order:
                    valid_groups = set(merged_data['Group'].unique())
                    ordered_groups = [group for group in self.group_order if group in valid_groups]
                    remaining_groups = sorted(list(valid_groups - set(ordered_groups)))
                    unique_groups = ordered_groups + remaining_groups
                else:
                    unique_groups = sorted(merged_data['Group'].unique())

                # Add replicate numbers
                merged_data['Replicate'] = merged_data.groupby(['Sample', 'Group']).cumcount()
                
                # Create pivot table
                result = pd.pivot_table(
                    merged_data,
                    values=selected_measurement,
                    index='Sample',
                    columns=['Group', 'Replicate'],
                    aggfunc='first'
                )
                
                # Reorder columns to keep replicates together within each group
                max_replicates = merged_data['Replicate'].max() + 1
                new_columns = []
                for group in unique_groups:
                    for rep in range(max_replicates):
                        if (group, rep) in result.columns:
                            new_columns.append((group, rep))
                
                # Reorder the columns and flatten the column names
                result = result.reindex(columns=new_columns)
                result.columns = [f"{col[0]}" for col in result.columns]
                result.index.name = None
                
            else:  # Mean & SD/SEM
                grouped = merged_data.groupby(['Sample', 'Group'])[selected_measurement]
                means = grouped.mean().round(2)
            
                if self.sd_radio.isChecked():
                    errors = grouped.std().round(2)
                    error_label = "SD"
                else:
                    with warnings.catch_warnings():
                        warnings.simplefilter("ignore")
                        errors = grouped.agg(lambda x: stats.sem(x, nan_policy='omit')).round(2)
            
                # Get unique groups in the specified order if available
                if self.group_order:
                    valid_groups = set(merged_data['Group'].unique())
                    ordered_groups = [group for group in self.group_order if group in valid_groups]
                    remaining_groups = sorted(list(valid_groups - set(ordered_groups)))
                    unique_groups = ordered_groups + remaining_groups
                else:
                    unique_groups = sorted(merged_data['Group'].unique())

                new_columns = []
                new_data = {}
            
                for group in unique_groups:
                    mean_col = f"{group}_Mean"
                    error_col = f"{group}_{error_label}"
                    new_columns.extend([mean_col, error_col])
                
                    group_means = means[means.index.get_level_values('Group') == group]
                    group_errors = errors[errors.index.get_level_values('Group') == group]
                
                    new_data[mean_col] = {idx[0]: val for idx, val in group_means.items()}
                    new_data[error_col] = {idx[0]: val for idx, val in group_errors.items()}
            
                result = pd.DataFrame(new_data, columns=new_columns)
                result.index.name = None

            # Apply sample order if available
            if self.sample_order:
                valid_order = [sample for sample in self.sample_order if sample in result.index]
                result = result.reindex(valid_order)

            return result

        except Exception as e:
            print(f"Error details: {str(e)}")
            print("Traceback:", traceback.format_exc())
            QMessageBox.critical(self, "Error", f"Error processing data: {str(e)}")
            return None

    def create_data_display_section(self):
        group_box = QGroupBox("Data Analysis")
        layout = QVBoxLayout()
        
        # Measurement selection
        measure_layout = QHBoxLayout()
        measure_layout.addWidget(QLabel("Select Measurement:"))
        self.measurement_combo = QComboBox()
        self.measurement_combo.setEnabled(False)
        measure_layout.addWidget(self.measurement_combo)
        layout.addLayout(measure_layout)
        
        # Filter options
        filter_group = QGroupBox("Filter Data")
        filter_layout = QHBoxLayout()
        self.filter_type_group = QButtonGroup()
        self.sample_radio = QRadioButton("Filter by Samples")
        self.group_radio = QRadioButton("Filter by Groups")
        self.sample_radio.setChecked(True)
        self.filter_type_group.addButton(self.sample_radio)
        self.filter_type_group.addButton(self.group_radio)
        filter_layout.addWidget(self.sample_radio)
        filter_layout.addWidget(self.group_radio)
        filter_group.setLayout(filter_layout)
        layout.addWidget(filter_group)
        
        self.sample_radio.toggled.connect(self.update_filter_list)
        self.group_radio.toggled.connect(self.update_filter_list)
    
        # Filter list
        self.filter_list = QListWidget()
        self.filter_list.setSelectionMode(QListWidget.SelectionMode.ExtendedSelection)
        self.filter_list.addItem("All")
        layout.addWidget(self.filter_list)
        
        group_box.setLayout(layout)
        return group_box
        
    def create_analysis_options_section(self):
        group_box = QGroupBox("Analysis Options")
        layout = QHBoxLayout()
        
        self.output_type_group = QButtonGroup()
        self.individual_radio = QRadioButton("Individual Replicates")
        self.sd_radio = QRadioButton("Mean & SD")
        self.sem_radio = QRadioButton("Mean & SEM")
        
        self.individual_radio.setChecked(True)
        self.output_type_group.addButton(self.individual_radio)
        self.output_type_group.addButton(self.sd_radio)
        self.output_type_group.addButton(self.sem_radio)
        
        layout.addWidget(self.individual_radio)
        layout.addWidget(self.sd_radio)
        layout.addWidget(self.sem_radio)
        layout.addStretch()
        
        group_box.setLayout(layout)
        return group_box

    def update_filter_list(self):
        # Clear current items
        self.filter_list.clear()
    
        # Add "All" option
        self.filter_list.addItem("All")
    
        # Add other options based on selected filter type
        if self.sample_radio.isChecked() and self.sample_map is not None:
            if self.sample_order:
                # Use sample order if available
                valid_samples = set(self.sample_map['Value'].unique())
                ordered_samples = [sample for sample in self.sample_order if sample in valid_samples]
                
                # Add any remaining samples that weren't in the order list
                remaining_samples = sorted(list(valid_samples - set(ordered_samples)))
                all_samples = ordered_samples + remaining_samples
                
                for sample in all_samples:
                    self.filter_list.addItem(str(sample))
            else:
                # Default to alphabetical order if no sample order provided
                options = sorted(self.sample_map['Value'].unique().tolist())
                for option in options:
                    self.filter_list.addItem(str(option))
                    
        elif self.group_radio.isChecked() and self.group_map is not None:
            if self.group_order:
                # Use group order if available
                valid_groups = set(self.group_map['Value'].unique())
                ordered_groups = [group for group in self.group_order if group in valid_groups]
                
                # Add any remaining groups that weren't in the order list
                remaining_groups = sorted(list(valid_groups - set(ordered_groups)))
                all_groups = ordered_groups + remaining_groups
                
                for group in all_groups:
                    self.filter_list.addItem(str(group))
            else:
                # Default to alphabetical order if no group order provided
                options = sorted(self.group_map['Value'].unique().tolist())
                for option in options:
                    self.filter_list.addItem(str(option))
            
        # Select "All" by default
        self.filter_list.setCurrentRow(0)
    
    def create_export_section(self):
        group_box = QGroupBox("Export Options")
        layout = QVBoxLayout()
        
        # Format options
        format_layout = QHBoxLayout()
        self.format_group = QButtonGroup()
        self.standard_radio = QRadioButton("Standard Format")
        self.single_row_radio = QRadioButton("Single Row Format")
        self.standard_radio.setChecked(True)
        self.format_group.addButton(self.standard_radio)
        self.format_group.addButton(self.single_row_radio)
        
        self.include_header_check = QCheckBox("Include Header")
        self.include_header_check.setChecked(True)
        
        format_layout.addWidget(self.standard_radio)
        format_layout.addWidget(self.single_row_radio)
        format_layout.addWidget(self.include_header_check)
        format_layout.addStretch()
        layout.addLayout(format_layout)
        
        # Export buttons
        button_layout = QHBoxLayout()
        self.copy_btn = QPushButton("Copy to Clipboard")
        self.save_btn = QPushButton("Save to CSV")
        self.copy_btn.clicked.connect(self.copy_to_clipboard)
        self.save_btn.clicked.connect(self.save_to_csv)
        
        button_layout.addWidget(self.copy_btn)
        button_layout.addWidget(self.save_btn)
        button_layout.addStretch()
        layout.addLayout(button_layout)
        
        group_box.setLayout(layout)
        return group_box

    def update_ui_state(self):
        all_loaded = all([self.sample_map is not None,
                         self.group_map is not None,
                         self.flow_data is not None])
        
        self.measurement_combo.setEnabled(all_loaded)
        self.filter_list.setEnabled(all_loaded)
        self.individual_radio.setEnabled(all_loaded)
        self.sd_radio.setEnabled(all_loaded)
        self.sem_radio.setEnabled(all_loaded)
        self.copy_btn.setEnabled(all_loaded)
        self.save_btn.setEnabled(all_loaded)
        

    def view_sample_plate(self):
        WellPlateViewer(self.root, "Sample Map Plate View", self.sample_well_data)

    def view_group_plate(self):
        WellPlateViewer(self.root, "Group Map Plate View", self.group_well_data)

    def process_plate_map(self, df):
        # Get well positions and values from the plate map
        well_data = []
        well_dict = {}
    
        # Print shape and contents for debugging
        print(f"DataFrame shape: {df.shape}")
        print("DataFrame contents:")
        print(df)
    
        # Find the last non-empty row in the DataFrame
        last_row = 0
        for idx in range(len(df.index) - 1, -1, -1):  # Include all rows
            row_data = df.iloc[idx].iloc[1:]  # Skip first column (row labels)
            if not row_data.isna().all():  # Check if row has any non-NA values
                last_row = idx
                print(f"Found non-empty row at index {idx}")
                break
    
        print(f"Last non-empty row index: {last_row}")
    
        # Process rows from the beginning (including first row)
        for row_idx in range(last_row + 1):  # Include all rows up to last_row
            row_letter = chr(65 + row_idx)  # Convert to A-H (no adjustment needed)
            row_data = df.iloc[row_idx]
            num_cols = len(row_data)  # Get total number of columns
        
            print(f"\nProcessing row {row_idx+1} (Row {row_letter}):")
            print(f"Row data: {row_data.values}")
        
            # Process each column, starting from column 1 (index 1)
            for col_idx in range(1, num_cols):
                try:
                    value = row_data.iloc[col_idx]
                    if pd.notna(value):  # Only process non-empty wells
                        well_position = f"{row_letter}{col_idx:02d}"  # Format as A01, A02, etc.
                        well_data.append({
                            'Well': well_position,
                            'Value': value
                        })
                        well_dict[well_position] = value
                        print(f"Processed well {well_position}: {value}")  # Debug output
                except IndexError as e:
                    print(f"IndexError at row {row_letter}, column {col_idx}: {e}")
                    continue
    
        print("\nProcessed wells summary:")
        for well in sorted(well_dict.keys()):
            print(f"{well}: {well_dict[well]}")
        
        return pd.DataFrame(well_data), well_dict
                
    def load_group_map(self):
        filename, _ = QFileDialog.getOpenFileName(self, "Load Group Map",
                                                "", "Excel Files (*.xlsx *.xls)")
        if filename:
            try:
                # Read the entire Excel file
                df = pd.read_excel(filename)
                
                # Extract group order from column 14 (index 13), starting from row 1 (ignoring the header)
                self.group_order = df.iloc[0:, 13].dropna().tolist()
                print(f"Loaded group order from column 14: {self.group_order}")
                
                # Process the plate map as before
                self.group_map, self.group_well_data = self.process_plate_map(df)
                self.group_status.setText("LOADED")
                self.group_status.setStyleSheet("color: green")
            
                # Update group options
                if self.group_radio.isChecked():
                    self.update_filter_list()
            
                self.update_ui_state()
                
            except Exception as e:
                self.group_status.setText("LOAD ERROR")
                self.group_status.setStyleSheet("color: red")
                QMessageBox.critical(self, "Error", f"Error loading Group Map: {str(e)}")
            
    def load_flowjo_data(self):
        filename, _ = QFileDialog.getOpenFileName(self, "Load Flowjo data",
                                                "", "Excel Files (*.xlsx *.xls)")
        if filename:
            try:
                # Read the Excel file
                self.flow_data = pd.read_excel(filename)
        
                # Rename the first column to "Sample Name"
                self.flow_data = self.flow_data.rename(columns={self.flow_data.columns[0]: "Sample Name"})
                
                # Filter out "Mean" and "SD" rows
                self.flow_data = self.flow_data[~self.flow_data['Sample Name'].isin(['Mean', 'SD'])]
        
                # Convert percentage columns to float
                for col in self.flow_data.columns:
                    if self.flow_data[col].astype(str).str.contains('%').any():
                        # Remove '%' and convert to float
                        self.flow_data[col] = self.flow_data[col].astype(str).str.replace('%', '').astype(float)
                    
                # Extract plate positions from sample names
                self.flow_data['Well'] = self.flow_data['Sample Name'].str.extract(r'_([A-H]\d{2})\.fcs$')
        
                # Update measurement combo box
                measurements = self.flow_data.columns[1:-1].tolist()  # Exclude Sample Name and Well columns
                self.measurement_combo.clear()  # Clear existing items
                self.measurement_combo.addItems(measurements)  # Add new items
                self.measurement_combo.setEnabled(True)  # Enable the combo box
                if measurements:
                    self.measurement_combo.setCurrentIndex(0)  # Set first item as current
        
                print("\nFlowjo Data Structure:")
                print(self.flow_data.head())
                print("\nColumn Types:")
                print(self.flow_data.dtypes)
            
                self.flow_status.setText("LOADED")
                self.flow_status.setStyleSheet("color: green")
                self.update_ui_state()
            
            except Exception as e:
                self.flow_status.setText("LOAD ERROR")
                self.flow_status.setStyleSheet("color: red")
                QMessageBox.critical(self, "Error", f"Error loading Flowjo data: {str(e)}")
            
    def check_all_files_loaded(self):
        # Enable/disable analysis options based on whether all files are loaded
        all_loaded = all([self.sample_map is not None, 
                         self.group_map is not None, 
                         self.flow_data is not None])
        state = 'normal' if all_loaded else 'disabled'
    
        # Update states of widgets
        for child in self.analysis_frame.winfo_children():
            child.configure(state=state)
    
        # Enable/disable the measurement combo and filter controls
        self.measurement_combo['state'] = 'readonly' if all_loaded else 'disabled'
        self.filter_listbox.configure(state='normal' if all_loaded else 'disabled')
    
        # Enable/disable radio buttons
        for child in self.data_frame.winfo_children():
            if isinstance(child, ttk.Radiobutton):
                child.configure(state=state)
        
    def reshape_to_single_row(self, df):
        """
        Reshape data to single row format with combined Sample_Group headers using pandas melt.
        Keeps replicates together and adds empty values for missing replicates.
        """
        try:
            # Handle different DataFrame structures
            if isinstance(df.index, pd.MultiIndex):
                print("MultiIndex detected - maintaining current format")
                return df
            
            # Reset index and explicitly set the name of the index column
            df_reset = df.copy()
            df_reset.index.name = 'Sample'
            df_reset = df_reset.reset_index()
        
            # Melt the DataFrame
            melted = pd.melt(df_reset, id_vars=['Sample'], var_name='Group', value_name='Value')
        
            # Create combined sample_group column
            melted['sample_group'] = melted['Sample'] + '_' + melted['Group']
        
            # Count occurrences of each sample_group combination
            value_counts = melted['sample_group'].value_counts()
            max_replicates = value_counts.max()
        
            # Create a new DataFrame with ordered columns and proper spacing for replicates
            new_data = []
            new_columns = []
        
            # Process each unique sample_group in sorted order
            for sample_group in sorted(melted['sample_group'].unique()):
                # Get values for this sample_group
                values = melted[melted['sample_group'] == sample_group]['Value'].values
            
                # Add the values and extend with None for missing replicates
                for i in range(max_replicates):
                    new_columns.append(sample_group)
                    if i < len(values):
                        new_data.append(values[i])
                    else:
                        new_data.append(None)
        
            # Create result DataFrame
            result = pd.DataFrame([new_data], columns=new_columns)
        
            return result
        
        except Exception as e:
            print(f"Error reshaping data: {str(e)}")
            print("\nDataFrame info:")
            print(df.info())
            print("\nDataFrame columns:", df.columns)
            print("DataFrame index:", df.index)
            traceback.print_exc()
            raise
        
    def copy_to_clipboard(self):
        try:
            result = self.process_data()
            if result is not None:
                # Get format choice
                format_choice = "single_row" if self.single_row_radio.isChecked() else "standard"
                include_header = self.include_header_check.isChecked()
        
                if format_choice == "single_row":
                    print("\nConverting to single row format")
                    print("Original data:")
                    print(result)
                    # Reshape data to single row format
                    result = self.reshape_to_single_row(result)
                    include_index = False  # Never include index for single row format
                else:
                    # For standard format, keep the index (sample names)
                    include_index = True
        
                # Convert to string with specific options for clean output
                result_str = result.to_csv(
                    sep='\t',
                    index=include_index,  # Include index for standard format
                    header=include_header
                )
        
                # Remove any empty lines that might have been added
                result_str = '\n'.join(line for line in result_str.split('\n') if line.strip())
                pyperclip.copy(result_str)
        
        except Exception as e:
            print(f"Error details: {str(e)}")
            print("Traceback:", traceback.format_exc())
            QMessageBox.critical(self, "Error", f"Error copying to clipboard: {str(e)}")
        
    def save_to_csv(self):
        try:
            result = self.process_data()
            if result is not None:
                filename, _ = QFileDialog.getSaveFileName(
                    self,
                    "Save CSV File",
                    "",
                    "CSV files (*.csv)"
                )
                if filename:
                    if not filename.endswith('.csv'):
                        filename += '.csv'
                    
                    if self.single_row_radio.isChecked() and not isinstance(result.index, pd.MultiIndex):
                        result = self.reshape_to_single_row(result)
                    
                    result.to_csv(
                        filename,
                        index=isinstance(result.index, pd.MultiIndex),
                        header=self.include_header_check.isChecked()
                    )
                    QMessageBox.information(self, "Success", "Data saved to CSV")
                
        except Exception as e:
            print(f"Error details: {str(e)}")
            QMessageBox.critical(self, "Error", f"Error saving to CSV: {str(e)}")

def main():
    app = QApplication(sys.argv)
    window = FlowDataProcessor()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
