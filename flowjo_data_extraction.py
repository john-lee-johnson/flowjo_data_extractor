import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import numpy as np
from scipy import stats
import pyperclip
import re
import traceback
import sys
 
class FlowDataProcessor:
    def __init__(self, root):
        self.root = root
        self.root.title("Flow Cytometry Data Processor")
        
        # Data storage
        self.sample_map = None
        self.group_map = None
        self.flow_data = None
        self.processed_data = None
        self.sample_well_data = {}
        self.group_well_data = {}
        
        # Create main frame
        self.main_frame = ttk.Frame(root, padding="10")
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure column weights
        self.main_frame.columnconfigure(0, weight=1)
        
        # File loading section
        self.create_file_loading_section()
        
        # Data display section
        self.create_data_display_section()
        
        # Analysis options section
        self.create_analysis_options_section()
        
        # Export section
        self.create_export_section()

    def create_file_loading_section(self):
        # File loading frame
        load_frame = ttk.LabelFrame(self.main_frame, text="Load Files", padding="5")
        load_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        load_frame.columnconfigure(1, weight=1)
        
        # Sample Map row
        sample_frame = ttk.Frame(load_frame)
        sample_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=2)
        ttk.Button(sample_frame, text="Load Sample Map", 
                  command=self.load_sample_map).pack(side=tk.LEFT, padx=5)
        self.sample_status = ttk.Label(sample_frame, text="No Sample Map", foreground="red")
        self.sample_status.pack(side=tk.LEFT, padx=5)
        
        # Group Map row
        group_frame = ttk.Frame(load_frame)
        group_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=2)
        ttk.Button(group_frame, text="Load Group Map", 
                  command=self.load_group_map).pack(side=tk.LEFT, padx=5)
        self.group_status = ttk.Label(group_frame, text="No Group Map", foreground="red")
        self.group_status.pack(side=tk.LEFT, padx=5)
        
        # Flowjo Data row
        flow_frame = ttk.Frame(load_frame)
        flow_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=2)
        ttk.Button(flow_frame, text="Load Flowjo Data", 
                  command=self.load_flowjo_data).pack(side=tk.LEFT, padx=5)
        self.flow_status = ttk.Label(flow_frame, text="No Flowjo Data", foreground="red")
        self.flow_status.pack(side=tk.LEFT, padx=5)

    def create_data_display_section(self):
        # Create frames for displaying loaded data
        self.data_frame = ttk.LabelFrame(self.main_frame, text="Data Analysis", padding="5")
        self.data_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        self.data_frame.columnconfigure(1, weight=1)

        # Measurement selection
        ttk.Label(self.data_frame, text="Select Measurement:").grid(row=0, column=0, padx=5, pady=5)
        self.measurement_var = tk.StringVar()
        self.measurement_combo = ttk.Combobox(self.data_frame, textvariable=self.measurement_var, 
                                            state='disabled', width=60)
        self.measurement_combo.grid(row=0, column=1, columnspan=2, padx=5, pady=5, sticky=tk.W+tk.E)

        # Create filter frame
        filter_frame = ttk.LabelFrame(self.data_frame, text="Filter Data")
        filter_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5, padx=5)

        # Radio buttons for selection type
        self.filter_type = tk.StringVar(value="sample")
        ttk.Radiobutton(filter_frame, text="Filter by Samples", 
                        variable=self.filter_type, 
                        value="sample",
                        command=self.update_filter_display).pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(filter_frame, text="Filter by Groups", 
                        variable=self.filter_type, 
                        value="group",
                        command=self.update_filter_display).pack(side=tk.LEFT, padx=5)

        # Create selection frame
        selection_frame = ttk.Frame(self.data_frame)
        selection_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        selection_frame.columnconfigure(0, weight=1)

        # Available options display
        available_frame = ttk.LabelFrame(selection_frame, text="Available Options")
        available_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Create and configure listbox with extended selection mode
        self.filter_listbox = tk.Listbox(available_frame, 
                                        selectmode=tk.EXTENDED, 
                                        height=6,
                                        exportselection=False)
        self.filter_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Configure scrollbar
        scrollbar = ttk.Scrollbar(available_frame, orient=tk.VERTICAL)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Link scrollbar and listbox
        self.filter_listbox.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.filter_listbox.yview)

        # Initialize with "All" option
        self.filter_listbox.insert(tk.END, "All")
        self.filter_listbox.select_set(0)

        # Store options
        self.sample_options = ["All"]
        self.group_options = ["All"]

        # Store last clicked index for shift-click functionality
        self.last_clicked_index = 0

        # Bind mouse events with platform-specific modifications
        if sys.platform == 'darwin':  # macOS
            self.filter_listbox.bind('<Button-1>', self.on_listbox_click)
            self.filter_listbox.bind('<Command-Button-1>', self.on_listbox_cmd_click)
        else:  # Windows/Linux
            self.filter_listbox.bind('<Button-1>', self.on_listbox_click)
            self.filter_listbox.bind('<Control-Button-1>', self.on_listbox_cmd_click)
    
        self.filter_listbox.bind('<Shift-Button-1>', self.on_listbox_shift_click)
        self.filter_listbox.bind('<B1-Motion>', self.on_listbox_drag)

    def on_listbox_click(self, event):
        """Handle single click selection"""
        index = self.filter_listbox.nearest(event.y)
        if index >= 0:
            # Check for Command key on Mac
            if sys.platform == 'darwin':
                has_modifier = event.state & 0x08  # Command key
            else:
                has_modifier = event.state & 0x04  # Control key
            
            if not has_modifier and not (event.state & 0x01):  # No Shift key
                self.filter_listbox.selection_clear(0, tk.END)
            self.filter_listbox.selection_set(index)
            self.filter_listbox.activate(index)
            self.last_clicked_index = index

    def on_listbox_cmd_click(self, event):
        """Handle Command/Control-click selection (toggle individual item)"""
        index = self.filter_listbox.nearest(event.y)
        if index >= 0:
            if self.filter_listbox.selection_includes(index):
                self.filter_listbox.selection_clear(index)
            else:
                self.filter_listbox.selection_set(index)
            self.filter_listbox.activate(index)
            self.last_clicked_index = index

    def on_listbox_shift_click(self, event):
        """Handle Shift-click selection (select range)"""
        index = self.filter_listbox.nearest(event.y)
        if index >= 0:
            # Select all items between last clicked and current
            start = min(self.last_clicked_index, index)
            end = max(self.last_clicked_index, index)
            self.filter_listbox.selection_clear(0, tk.END)
            for i in range(start, end + 1):
                self.filter_listbox.selection_set(i)
            self.filter_listbox.activate(index)

    def on_listbox_drag(self, event):
        """Handle drag selection"""
        index = self.filter_listbox.nearest(event.y)
        if index >= 0:
            # If shift or cmd/ctrl is not held, clear other selections
            if not (event.state & 0x0001) and not (event.state & 0x0004):
                self.filter_listbox.selection_clear(0, tk.END)
            self.filter_listbox.selection_set(index)
            self.filter_listbox.activate(index)
            
    def update_filter_display(self):
        """Update the listbox based on the selected filter type"""
        # Store current selection focus
        current_selection = self.filter_listbox.curselection()
        focused_index = self.filter_listbox.index(tk.ACTIVE) if current_selection else 0
    
        # Clear and repopulate the listbox
        self.filter_listbox.delete(0, tk.END)
    
        # Insert new options without any processing
        options = self.sample_options if self.filter_type.get() == "sample" else self.group_options
        for option in options:
            self.filter_listbox.insert(tk.END, option)

        # Restore selection state
        if current_selection:
            try:
                self.filter_listbox.select_set(focused_index)
                self.filter_listbox.activate(focused_index)
            except tk.TclError:
                # If the previous index is no longer valid, select the first item
                self.filter_listbox.select_set(0)
                self.filter_listbox.activate(0)
        else:
            self.filter_listbox.select_set(0)
            self.filter_listbox.activate(0)
        
    def create_analysis_options_section(self):
        # Analysis options
        self.analysis_frame = ttk.LabelFrame(self.main_frame, text="Analysis Options", padding="5")
        self.analysis_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        
        self.output_type = tk.StringVar(value="individual")
        ttk.Radiobutton(self.analysis_frame, text="Individual Replicates", 
                       variable=self.output_type, value="individual").grid(row=0, column=0, padx=5, pady=5)
        ttk.Radiobutton(self.analysis_frame, text="Mean & SD", 
                       variable=self.output_type, value="sd").grid(row=0, column=1, padx=5, pady=5)
        ttk.Radiobutton(self.analysis_frame, text="Mean & SEM", 
                       variable=self.output_type, value="sem").grid(row=0, column=2, padx=5, pady=5)

    def create_export_section(self):
        try:
            # Export frame
            export_frame = ttk.LabelFrame(self.main_frame, text="Export Options", padding="5")
            export_frame.grid(row=3, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        
            # Create frame for format options
            format_frame = ttk.Frame(export_frame)
            format_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
            # Add format radio buttons
            self.format_var = tk.StringVar(value="standard")
            ttk.Radiobutton(format_frame, text="Standard Format", 
                           variable=self.format_var, value="standard").grid(row=0, column=0, padx=5)
            ttk.Radiobutton(format_frame, text="Single Row Format", 
                           variable=self.format_var, value="single_row").grid(row=0, column=1, padx=5)
        
            # Add header checkbox
            self.include_header = tk.BooleanVar(value=True)
            ttk.Checkbutton(format_frame, text="Include Header", 
                           variable=self.include_header).grid(row=0, column=2, padx=5)
        
            # Add export buttons
            button_frame = ttk.Frame(export_frame)
            button_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
            ttk.Button(button_frame, text="Copy to Clipboard", 
                      command=self.copy_to_clipboard).grid(row=0, column=0, padx=5)
            ttk.Button(button_frame, text="Save to CSV", 
                      command=self.save_to_csv).grid(row=0, column=1, padx=5)
        except Exception as e:
            print(f"Error creating export section: {str(e)}")
            raise

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
    
    def load_sample_map(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if filename:
            try:
                df = pd.read_excel(filename)
                self.sample_map, self.sample_well_data = self.process_plate_map(df)
                self.sample_status.config(text="LOADED", foreground="green")
            
                # Update sample options
                self.sample_options = ["All"] + sorted(self.sample_map['Value'].unique().tolist())
            
                # Update display if currently showing samples
                if self.filter_type.get() == "sample":
                    self.update_filter_display()
            
                self.check_all_files_loaded()
            except Exception as e:
                self.sample_status.config(text="LOAD ERROR", foreground="red")
                messagebox.showerror("Error", f"Error loading Sample Map: {str(e)}")

    def load_group_map(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if filename:
            try:
                df = pd.read_excel(filename)
                self.group_map, self.group_well_data = self.process_plate_map(df)
                self.group_status.config(text="LOADED", foreground="green")
            
                # Update group options
                self.group_options = ["All"] + sorted(self.group_map['Value'].unique().tolist())
            
                # Update display if currently showing groups
                if self.filter_type.get() == "group":
                    self.update_filter_display()
            
                self.check_all_files_loaded()
            except Exception as e:
                self.group_status.config(text="LOAD ERROR", foreground="red")
                messagebox.showerror("Error", f"Error loading Group Map: {str(e)}")

    def load_flowjo_data(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if filename:
            try:
                # Read the Excel file
                self.flow_data = pd.read_excel(filename)
            
                # Rename the first column to "Sample Name"
                self.flow_data = self.flow_data.rename(columns={self.flow_data.columns[0]: "Sample Name"})
            
                # Convert percentage columns to float
                for col in self.flow_data.columns:
                    if self.flow_data[col].astype(str).str.contains('%').any():
                        # Remove '%' and convert to float
                        self.flow_data[col] = self.flow_data[col].astype(str).str.replace('%', '').astype(float)
                        
            
                # Extract plate positions from sample names
                self.flow_data['Well'] = self.flow_data['Sample Name'].str.extract(r'_([A-H]\d{2})\.fcs$')
            
                # Update measurement combo box
                measurements = self.flow_data.columns[1:-1].tolist()  # Exclude Sample Name and Well columns
                self.measurement_combo['values'] = measurements
                self.measurement_combo['state'] = 'readonly'
                if measurements:
                    self.measurement_combo.set(measurements[0])
            
                print("\nFlowjo Data Structure:")
                print(self.flow_data.head())
                print("\nColumn Types:")
                print(self.flow_data.dtypes)
            
                self.flow_status.config(text="LOADED", foreground="green")
                self.check_all_files_loaded()
            except Exception as e:
                self.flow_status.config(text="LOAD ERROR", foreground="red")
                messagebox.showerror("Error", f"Error loading Flowjo Data: {str(e)}")
            
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

    def process_data(self):
        if not all([self.sample_map is not None, self.group_map is not None, self.flow_data is not None]):
            messagebox.showerror("Error", "Please load all required files first")
            return None

        selected_measurement = self.measurement_var.get()
        if not selected_measurement:
            messagebox.showerror("Error", "Please select a measurement")
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
        
            # Get selected options
            selected_indices = self.filter_listbox.curselection()
            if not selected_indices:  # If nothing is selected, assume "All"
                selected_options = ["All"]
            else:
                selected_options = [self.filter_listbox.get(idx) for idx in selected_indices]
        
            # Apply filters based on selection type
            if self.filter_type.get() == "sample" and "All" not in selected_options:
                merged_data = merged_data[merged_data['Sample'].isin(selected_options)]
            elif self.filter_type.get() == "group" and "All" not in selected_options:
                merged_data = merged_data[merged_data['Group'].isin(selected_options)]
            
            # Ensure numeric type for selected measurement
            merged_data[selected_measurement] = pd.to_numeric(merged_data[selected_measurement], errors='coerce')
        
            # Process based on selected output type
            output_type = self.output_type.get()
            if output_type == "individual":
                # Use cumcount to number the replicates within each Sample-Group combination
                merged_data['Replicate'] = merged_data.groupby(['Sample', 'Group']).cumcount()
            
                # Pivot the data to get replicates as columns
                result = pd.pivot_table(
                    merged_data,
                    values=selected_measurement,
                    index='Sample',
                    columns=['Group', 'Replicate'],
                    aggfunc='first'  # Each value should appear only once
                )
            
                # Flatten column names to just use the group name
                result.columns = [f"{col[0]}" for col in result.columns]
            
                # Sort columns to keep groups together
                result = result.reindex(sorted(result.columns), axis=1)
            
                # Set index name to None to remove "Sample" row
                result.index.name = None
            
            else:  # Mean & SD/SEM
                grouped = merged_data.groupby(['Sample', 'Group'])[selected_measurement]
                means = grouped.mean().unstack()
                means.index.name = None
            
                if output_type == "sd":
                    errors = grouped.std().unstack()
                else:  # SEM
                    errors = grouped.agg(lambda x: stats.sem(x, nan_policy='omit')).unstack()
            
                errors.index.name = None
                result = pd.concat([means, errors], keys=['Mean', 'Error'])

            return result

        except Exception as e:
            print(f"Error details: {str(e)}")
            print("Traceback:", traceback.format_exc())
            messagebox.showerror("Error", f"Error processing data: {str(e)}")
            return None
    
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
                format_choice = self.format_var.get()
                include_header = self.include_header.get()
            
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
            messagebox.showerror("Error", f"Error copying to clipboard: {str(e)}")
        
    def save_to_csv(self):
        try:
            result = self.process_data()
            if result is not None:
                filename = filedialog.asksaveasfilename(
                    defaultextension=".csv",
                    filetypes=[("CSV files", "*.csv")]
                )
                if filename:
                    if self.format_var.get() == "single_row" and not isinstance(result.index, pd.MultiIndex):
                        result = self.reshape_to_single_row(result)
                    result.to_csv(
                        filename,
                        index=isinstance(result.index, pd.MultiIndex),
                        header=self.include_header.get()
                    )
                    messagebox.showinfo("Success", "Data saved to CSV")
                
        except Exception as e:
            print(f"Error details: {str(e)}")
            messagebox.showerror("Error", f"Error saving to CSV: {str(e)}")

def main():
    root = tk.Tk()
    app = FlowDataProcessor(root)
    
    # Configure better event handling
    root.update_idletasks()
    root.after(100, root.update_idletasks)  # Schedule periodic updates
    
    root.mainloop()

if __name__ == "__main__":
    main()