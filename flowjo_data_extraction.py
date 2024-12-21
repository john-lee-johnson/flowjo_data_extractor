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
import os

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
        self.flow_data_files = []  # List to store multiple flow data DataFrames
        self.flow_data_names = []  # List to store the names of loaded files
    
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
        flow_layout = QVBoxLayout()
        flow_buttons_layout = QHBoxLayout()
        
        # Split into two buttons: Add and Remove
        self.add_flow_btn = QPushButton("Add Flow Data File")
        self.remove_flow_btn = QPushButton("Remove Selected")
        self.add_flow_btn.clicked.connect(self.add_flowjo_data)
        self.remove_flow_btn.clicked.connect(self.remove_flowjo_data)
        self.remove_flow_btn.setEnabled(False)  # Disabled until files are loaded
        
        flow_buttons_layout.addWidget(self.add_flow_btn)
        flow_buttons_layout.addWidget(self.remove_flow_btn)
        flow_buttons_layout.addStretch()
        
        self.flow_status = QLabel("No Flow Data")
        self.flow_status.setStyleSheet("color: red")
        self.flow_files_list = QListWidget()
        self.flow_files_list.itemSelectionChanged.connect(self.update_remove_button_state)
        
        flow_layout.addLayout(flow_buttons_layout)
        flow_layout.addWidget(self.flow_status)
        flow_layout.addWidget(self.flow_files_list)
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
                
                # Check if column 14 exists before trying to read from it
                if df.shape[1] > 13:  # Check if there are at least 14 columns
                    sample_order_col = df.iloc[0:, 13].dropna()
                    self.sample_order = sample_order_col.tolist() if not sample_order_col.empty else None
                else:
                    self.sample_order = None
                
                print(f"Loaded sample order from column 14: {self.sample_order if self.sample_order else 'None - using default order'}")
                
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
        if not all([self.sample_map is not None, 
                    self.group_map is not None, 
                    len(self.flow_data_files) > 0]):
            QMessageBox.critical(self, "Error", "Please load all required files first")
            return None

        selected_measurement = self.measurement_combo.currentText()
        if not selected_measurement:
            QMessageBox.critical(self, "Error", "Please select a measurement")
            return None

        try:
            # Process each flow data file
            results = []
            for flow_data in self.flow_data_files:
                # Merge flow data with sample information first
                merged_data = flow_data.merge(
                    self.sample_map,
                    on='Well',
                    how='left'
                )
                merged_data = merged_data.rename(columns={'Value': 'Sample'})
    
                # Check if Sample column is all NaN
                if merged_data['Sample'].isna().all():
                    QMessageBox.warning(
                        self,
                        "Warning",
                        "No matches found in Sample Map. Please check that well positions in your Sample Map match the data file."
                    )
                    return None

                # Then merge with group information
                merged_data = merged_data.merge(
                    self.group_map,
                    on='Well',
                    how='left'
                )
                merged_data = merged_data.rename(columns={'Value': 'Group'})

                # Check if Group column is all NaN
                if merged_data['Group'].isna().all():
                    QMessageBox.warning(
                        self,
                        "Warning",
                        "No matches found in Group Map. Please check that well positions in your Group Map match the data file."
                    )
                    return None
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
                        valid_groups = set(merged_data['Group'].astype(str).unique())  # Convert to string
                        ordered_groups = [group for group in self.group_order if str(group) in valid_groups]
                        remaining_groups = sorted(list(valid_groups - set(map(str, ordered_groups))))
                        unique_groups = ordered_groups + remaining_groups
                    else:
                        # Convert all groups to strings before sorting
                        unique_groups = sorted(merged_data['Group'].astype(str).unique())

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
                    
                    

                    # Convert max_replicates to integer
                    max_replicates = int(merged_data['Replicate'].max() + 1)
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
                        valid_groups = set(merged_data['Group'].astype(str).unique())  # Convert to string
                        ordered_groups = [group for group in self.group_order if str(group) in valid_groups]
                        remaining_groups = sorted(list(valid_groups - set(map(str, ordered_groups))))
                        unique_groups = ordered_groups + remaining_groups
                    else:
                        # Convert all groups to strings before sorting
                        unique_groups = sorted(merged_data['Group'].astype(str).unique())

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

                results.append(result)

            # If we're in XY row format, return all results
            if self.XY_radio.isChecked():
                return self.reshape_to_xy_format(results)
            else:
                # Otherwise, just return the first result
                return results[0] if results else None

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
        self.filter_list.clear()
        self.filter_list.addItem("All")

        if self.sample_radio.isChecked() and self.sample_map is not None:
            samples = self.sample_map['Value'].unique()
            if self.sample_order:
                # Use sample order if available
                valid_samples = set(samples)
                ordered_samples = [sample for sample in self.sample_order if sample in valid_samples]
                remaining_samples = sorted(list(valid_samples - set(ordered_samples)))
                all_samples = ordered_samples + remaining_samples
            else:
                # Use original order if no sample order provided
                all_samples = sorted(samples)
            
            for sample in all_samples:
                self.filter_list.addItem(str(sample))

        elif self.group_radio.isChecked() and self.group_map is not None:
            groups = self.group_map['Value'].unique()
            if self.group_order:
                # Use group order if available
                valid_groups = set(groups)
                ordered_groups = [group for group in self.group_order if group in valid_groups]
                remaining_groups = sorted(list(valid_groups - set(ordered_groups)))
                all_groups = ordered_groups + remaining_groups
            else:
                # Use original order if no group order provided
                all_groups = sorted(groups)
            
            for group in all_groups:
                self.filter_list.addItem(str(group))

        self.filter_list.setCurrentRow(0)
    
    def create_export_section(self):
        group_box = QGroupBox("Export Options")
        layout = QVBoxLayout()
        
        # Format options
        format_layout = QHBoxLayout()
        self.format_group = QButtonGroup()
        self.standard_radio = QRadioButton("Standard Format")
        self.XY_radio = QRadioButton("XY Format")
        self.standard_radio.setChecked(True)
        self.format_group.addButton(self.standard_radio)
        self.format_group.addButton(self.XY_radio)
        
        self.include_header_check = QCheckBox("Include Header")
        self.include_header_check.setChecked(True)
        
        format_layout.addWidget(self.standard_radio)
        format_layout.addWidget(self.XY_radio)
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
        # Update based on whether all required files are loaded
        all_loaded = all([
            self.sample_map is not None,
            self.group_map is not None,
            len(self.flow_data_files) > 0  # Changed from self.flow_data check
        ])
        
        self.measurement_combo.setEnabled(all_loaded)
        self.filter_list.setEnabled(all_loaded)
        self.individual_radio.setEnabled(all_loaded)
        self.sd_radio.setEnabled(all_loaded)
        self.sem_radio.setEnabled(all_loaded)
        self.copy_btn.setEnabled(all_loaded)
        self.save_btn.setEnabled(all_loaded)

    def update_measurement_combo(self):
        if not self.flow_data_files:
            self.measurement_combo.clear()
            self.measurement_combo.setEnabled(False)
            return
        
        # Find common columns across all files
        common_columns = set(self.flow_data_files[0].columns[1:-1])  # Exclude 'Sample Name' and 'Well'
        for flow_data in self.flow_data_files[1:]:
            common_columns &= set(flow_data.columns[1:-1])
        
        # Update combo box
        self.measurement_combo.clear()
        self.measurement_combo.addItems(sorted(list(common_columns)))
        self.measurement_combo.setEnabled(True)
        if common_columns:
            self.measurement_combo.setCurrentIndex(0)
        
        print(f"Common columns found: {common_columns}")  # Debug output

    def add_flowjo_data(self):
        filename, _ = QFileDialog.getOpenFileName(
            self,
            "Add Flow Data File",
            "",
            "Excel Files (*.xlsx *.xls)"
        )
        if filename:
            try:
                # Read and process the new file
                flow_data = pd.read_excel(filename)
                flow_data = flow_data.rename(columns={flow_data.columns[0]: "Sample Name"})
                flow_data = flow_data[~flow_data['Sample Name'].isin(['Mean', 'SD'])]
                
                for col in flow_data.columns:
                    if flow_data[col].astype(str).str.contains('%').any():
                        flow_data[col] = flow_data[col].astype(str).str.replace('%', '').astype(float)
                
                flow_data['Well'] = flow_data['Sample Name'].str.extract(r'_([A-H]\d{2})\.fcs$')
                
                # Add to lists
                self.flow_data_files.append(flow_data)
                file_name = os.path.basename(filename)
                self.flow_data_names.append(file_name)
                self.flow_files_list.addItem(file_name)
                
                # Update status
                self.flow_status.setText(f"LOADED ({len(self.flow_data_files)} files)")
                self.flow_status.setStyleSheet("color: green")
                
                # Update measurement combo box and UI state
                self.update_measurement_combo()
                self.update_ui_state()
                
                print(f"Added file: {file_name}")  # Debug output
                print(f"Total files: {len(self.flow_data_files)}")  # Debug output
                
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Error loading file: {str(e)}")
                print(f"Error loading file: {str(e)}")  # Debug output

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
                except IndexError as e:
                    print(f"IndexError at row {row_letter}, column {col_idx}: {e}")
                    continue
    
        return pd.DataFrame(well_data), well_dict
                
    def load_group_map(self):
        filename, _ = QFileDialog.getOpenFileName(self, "Load Group Map",
                                                "", "Excel Files (*.xlsx *.xls)")
        if filename:
            try:
                # Read the entire Excel file
                df = pd.read_excel(filename)
                
                # Check if column 14 exists before trying to read from it
                if df.shape[1] > 13:  # Check if there are at least 14 columns
                    group_order_col = df.iloc[0:, 13].dropna()
                    self.group_order = group_order_col.tolist() if not group_order_col.empty else None
                else:
                    self.group_order = None
                
                print(f"Loaded group order from column 14: {self.group_order if self.group_order else 'None - using default order'}")
                
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
        filenames, _ = QFileDialog.getOpenFileNames(
            self,
            "Load Flowjo data files",
            "",
            "Excel Files (*.xlsx *.xls)"
        )
        if filenames:
            try:
                self.flow_data_files = []
                self.flow_data_names = []
                self.flow_files_list.clear()
                
                for filename in filenames:
                    # Read the Excel file
                    flow_data = pd.read_excel(filename)
                    
                    # Process as before
                    flow_data = flow_data.rename(columns={flow_data.columns[0]: "Sample Name"})
                    flow_data = flow_data[~flow_data['Sample Name'].isin(['Mean', 'SD'])]
                    
                    for col in flow_data.columns:
                        if flow_data[col].astype(str).str.contains('%').any():
                            flow_data[col] = flow_data[col].astype(str).str.replace('%', '').astype(float)
                    
                    flow_data['Well'] = flow_data['Sample Name'].str.extract(r'_([A-H]\d{2})\.fcs$')
                    
                    self.flow_data_files.append(flow_data)
                    file_name = os.path.basename(filename)
                    self.flow_data_names.append(file_name)
                    self.flow_files_list.addItem(file_name)
                
                # Update measurement combo box using the first file
                if self.flow_data_files:
                    measurements = self.flow_data_files[0].columns[1:-1].tolist()
                    self.measurement_combo.clear()
                    self.measurement_combo.addItems(measurements)
                    self.measurement_combo.setEnabled(True)
                    if measurements:
                        self.measurement_combo.setCurrentIndex(0)
                
                self.flow_status.setText(f"LOADED ({len(self.flow_data_files)} files)")
                self.flow_status.setStyleSheet("color: green")
                self.update_ui_state()
                
            except Exception as e:
                self.flow_status.setText("LOAD ERROR")
                self.flow_status.setStyleSheet("color: red")
                QMessageBox.critical(self, "Error", f"Error loading Flowjo data: {str(e)}")
            
       
    def reshape_to_xy_format(self, df_list):
        """
        Reshape multiple DataFrames into XY format where:
        - Each row represents a different timepoint/data file
        - Columns are Sample_Group combinations
        - All rows have the same number of replicates (filled with None if needed)
        """
        try:
            all_columns = set()
            max_replicates_by_group = {}
            
            # First pass: collect all Sample_Group combinations and find max replicates
            for df in df_list:
                # Make a copy and reset index to create the Sample column
                df_reset = df.copy()
                df_reset.reset_index(inplace=True, names=['Sample'])
                
                # Melt the DataFrame
                melted = pd.melt(df_reset, 
                                id_vars=['Sample'],
                                var_name='Group', 
                                value_name='Value')
                
                # Combine sample and group
                melted['Sample_Group'] = melted['Sample'].astype(str) + '_' + melted['Group'].astype(str)
                
                # Add all Sample_Group combinations to the set
                all_columns.update(melted['Sample_Group'].unique())
                
                # Update max replicates for each Sample_Group
                for combined in melted['Sample_Group'].unique():
                    replicate_count = len(melted[melted['Sample_Group'] == combined])
                    max_replicates_by_group[combined] = max(
                        max_replicates_by_group.get(combined, 0),
                        replicate_count
                    )
            
            # Convert to sorted list for consistent column order
            all_columns = sorted(list(all_columns))
            
            # Second pass: create rows with consistent replicate counts
            all_rows = []
            for df in df_list:
                # Process each DataFrame
                df_reset = df.copy()
                df_reset.reset_index(inplace=True, names=['Sample'])
                
                melted = pd.melt(df_reset, 
                                id_vars=['Sample'],
                                var_name='Group', 
                                value_name='Value')
                
                melted['Sample_Group'] = melted['Sample'].astype(str) + '_' + melted['Group'].astype(str)
                
                # Create a row with values for all possible columns
                row_data = []
                for combined in all_columns:
                    values = melted[melted['Sample_Group'] == combined]['Value'].values
                    # Extend values with None to match max replicates for this group
                    values = list(values) + [None] * (max_replicates_by_group[combined] - len(values))
                    row_data.extend(values)
                
                all_rows.append(row_data)
            
            # Create column names with consistent replicate counts
            final_columns = []
            for combined in all_columns:
                final_columns.extend([combined] * max_replicates_by_group[combined])
            
            result = pd.DataFrame(all_rows, columns=final_columns)
            return result
        
        except Exception as e:
            print(f"Error reshaping data: {str(e)}")
            print("\nDataFrame list info:")
            for i, df in enumerate(df_list):
                print(f"\nDataFrame {i} info:")
                print(df.info())
                print(f"DataFrame {i} columns:", df.columns)
                print(f"DataFrame {i} index:", df.index)
            traceback.print_exc()
            raise
        
        
    def copy_to_clipboard(self):
        try:
            result = self.process_data()  # This already returns data in the correct format
            if result is not None:
                # Get format choice
                format_choice = "XY_format" if self.XY_radio.isChecked() else "standard"
                include_header = self.include_header_check.isChecked()
        
                if format_choice == "XY_format":
                    print("\nXY format selected")
                    print("Data:")
                    print(result)
                    include_index = False  # Never include index for XY format
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
                    
                    if self.XY_radio.isChecked() and not isinstance(result.index, pd.MultiIndex):
                        result = self.reshape_to_xy_format(result)
                    
                    result.to_csv(
                        filename,
                        index=isinstance(result.index, pd.MultiIndex),
                        header=self.include_header_check.isChecked()
                    )
                    QMessageBox.information(self, "Success", "Data saved to CSV")
                
        except Exception as e:
            print(f"Error details: {str(e)}")
            QMessageBox.critical(self, "Error", f"Error saving to CSV: {str(e)}")

    def remove_flowjo_data(self):
        # Get selected items
        selected_items = self.flow_files_list.selectedItems()
        
        for item in selected_items:
            # Get the index of the selected item
            index = self.flow_files_list.row(item)
            
            # Remove the file from both lists
            self.flow_data_files.pop(index)
            self.flow_data_names.pop(index)
            
            # Remove the item from the list widget
            self.flow_files_list.takeItem(self.flow_files_list.row(item))
        
        # Update the status
        if len(self.flow_data_files) > 0:
            self.flow_status.setText(f"LOADED ({len(self.flow_data_files)} files)")
            self.flow_status.setStyleSheet("color: green")
        else:
            self.flow_status.setText("No Flow Data")
            self.flow_status.setStyleSheet("color: red")
        
        # Update measurement combo box and UI state
        self.update_measurement_combo()
        self.update_ui_state()

    def update_remove_button_state(self):
        # Enable remove button only if there are selected items
        self.remove_flow_btn.setEnabled(len(self.flow_files_list.selectedItems()) > 0)

def main():
    app = QApplication(sys.argv)
    window = FlowDataProcessor()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
