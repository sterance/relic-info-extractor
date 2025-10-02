#!/usr/bin/env python3
"""
Tkinter GUI with Import/Export buttons and TreeView
A modern GUI application with file import/export capabilities and data display.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import json
import csv
import os
from typing import List, Dict, Any


class DataManagerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Relic Info Extractor")
        self.root.geometry("800x600")
        self.root.configure(bg='#2b2b2b')

        # Data storage
        self.data = []
        self.next_id = 1

        # Sorting state
        self.sort_column = None
        self.sort_reverse = False

        # Track previously used values for editable columns
        self.used_categories = set()
        self.used_display_groups = set()
        self.used_level_groups = set()
        self.used_levels = set()
        self.used_stacks = set()

        # Configure style
        self.setup_styles()

        # Create GUI components
        self.create_widgets()

    def setup_styles(self):
        """Configure dark modern styling for the GUI"""
        style = ttk.Style()
        style.theme_use('clam')

        # Configure dark color scheme
        style.configure('Dark.TFrame', background='#2b2b2b')
        style.configure('Dark.TLabel', background='#2b2b2b',
                        foreground='#ffffff')

        # Configure button styles with dark theme
        style.configure('Modern.TButton',
                        padding=(12, 8),
                        font=('Arial', 10, 'bold'),
                        background='#404040',
                        foreground='#ffffff',
                        borderwidth=0,
                        focuscolor='none')

        style.map('Modern.TButton',
                  background=[('active', '#505050'),
                              ('pressed', '#606060')],
                  foreground=[('active', '#ffffff')])

        # Configure treeview style with dark theme
        style.configure('Modern.Treeview',
                        font=('Arial', 9),
                        rowheight=28,
                        background='#3c3c3c',
                        foreground='#ffffff',
                        fieldbackground='#3c3c3c',
                        borderwidth=0)

        style.configure('Modern.Treeview.Heading',
                        font=('Arial', 10, 'bold'),
                        background='#404040',
                        foreground='#ffffff',
                        relief='flat')

        style.map('Modern.Treeview.Heading',
                  background=[('active', '#505050')])

        # Configure scrollbar style
        style.configure('Dark.Vertical.TScrollbar',
                        background='#404040',
                        troughcolor='#2b2b2b',
                        borderwidth=0,
                        arrowcolor='#ffffff',
                        darkcolor='#404040',
                        lightcolor='#404040')

        style.configure('Dark.Horizontal.TScrollbar',
                        background='#404040',
                        troughcolor='#2b2b2b',
                        borderwidth=0,
                        arrowcolor='#ffffff',
                        darkcolor='#404040',
                        lightcolor='#404040')

        # Configure status bar style
        style.configure('Dark.TLabel',
                        background='#404040',
                        foreground='#ffffff',
                        font=('Arial', 9))

    def create_widgets(self):
        """Create and layout all GUI widgets"""
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10", style='Dark.TFrame')
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(1, weight=1)

        # Button frame
        button_frame = ttk.Frame(main_frame, style='Dark.TFrame')
        button_frame.grid(row=1, column=0, sticky=(tk.W, tk.N), padx=(0, 10))

        # Import button
        self.import_btn = ttk.Button(button_frame,
                                     text="Import CSV",
                                     command=self.import_data,
                                     style='Modern.TButton')
        self.import_btn.grid(row=0, column=0, pady=(
            0, 10), sticky=(tk.W, tk.E))

        # Export button
        self.export_btn = ttk.Button(button_frame,
                                     text="Export JSON",
                                     command=self.export_data,
                                     style='Modern.TButton')
        self.export_btn.grid(row=1, column=0, pady=(
            0, 10), sticky=(tk.W, tk.E))

        # Clear button
        self.clear_btn = ttk.Button(button_frame,
                                    text="Clear Data",
                                    command=self.clear_data,
                                    style='Modern.TButton')
        self.clear_btn.grid(row=2, column=0, pady=(0, 10), sticky=(tk.W, tk.E))

        # Save Project button
        self.save_project_btn = ttk.Button(button_frame,
                                           text="Save Project",
                                           command=self.save_project,
                                           style='Modern.TButton')
        self.save_project_btn.grid(
            row=3, column=0, pady=(0, 10), sticky=(tk.W, tk.E))

        # Load Project button
        self.load_project_btn = ttk.Button(button_frame,
                                           text="Load Project",
                                           command=self.load_project,
                                           style='Modern.TButton')
        self.load_project_btn.grid(
            row=4, column=0, pady=(0, 10), sticky=(tk.W, tk.E))

        # Treeview frame
        tree_frame = ttk.Frame(main_frame, style='Dark.TFrame')
        tree_frame.grid(row=1, column=1, sticky=(tk.W, tk.E, tk.N, tk.S))
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)

        # Treeview with scrollbars
        self.tree = ttk.Treeview(tree_frame, style='Modern.Treeview')

        # Define columns for relic data
        self.tree['columns'] = ('id', 'gameIds', 'name', 'category', 'display_group',
                                'level_group_id', 'level_group', 'level', 'nightfarer', 'deep', 'debuff', 'stacks')
        self.tree.column('#0', width=0, stretch=False)
        self.tree.column('id', width=30, minwidth=20)
        self.tree.column('gameIds', width=200, minwidth=80)
        self.tree.column('name', width=550, minwidth=120)
        self.tree.column('category', width=100, minwidth=80)
        self.tree.column('display_group', width=200, minwidth=80)
        self.tree.column('level_group_id', width=100, minwidth=80)
        self.tree.column('level_group', width=200, minwidth=80)
        self.tree.column('level', width=60, minwidth=50)
        self.tree.column('nightfarer', width=100, minwidth=80)
        self.tree.column('deep', width=60, minwidth=50)
        self.tree.column('debuff', width=60, minwidth=50)
        self.tree.column('stacks', width=60, minwidth=50)

        # Define headings with sorting
        self.tree.heading('id', text='ID', anchor=tk.W,
                          command=lambda: self.sort_by_column('id'))
        self.tree.heading('gameIds', text='Game IDs', anchor=tk.W,
                          command=lambda: self.sort_by_column('gameIds'))
        self.tree.heading('name', text='Name', anchor=tk.W,
                          command=lambda: self.sort_by_column('name'))
        self.tree.heading('category', text='Category', anchor=tk.W,
                          command=lambda: self.sort_by_column('category'))
        self.tree.heading('display_group', text='Display Group', anchor=tk.W,
                          command=lambda: self.sort_by_column('display_group'))
        self.tree.heading('level_group_id', text='Level Group ID', anchor=tk.W,
                          command=lambda: self.sort_by_column('level_group_id'))
        self.tree.heading('level_group', text='Level Group', anchor=tk.W,
                          command=lambda: self.sort_by_column('level_group'))
        self.tree.heading('level', text='Level', anchor=tk.W,
                          command=lambda: self.sort_by_column('level'))
        self.tree.heading('nightfarer', text='Nightfarer', anchor=tk.W,
                          command=lambda: self.sort_by_column('nightfarer'))
        self.tree.heading('deep', text='Deep', anchor=tk.W,
                          command=lambda: self.sort_by_column('deep'))
        self.tree.heading('debuff', text='Debuff', anchor=tk.W,
                          command=lambda: self.sort_by_column('debuff'))
        self.tree.heading('stacks', text='Stacks', anchor=tk.W,
                          command=lambda: self.sort_by_column('stacks'))

        # Scrollbars
        v_scrollbar = ttk.Scrollbar(
            tree_frame, orient=tk.VERTICAL, command=self.tree.yview, style='Dark.Vertical.TScrollbar')

        # h_scrollbar = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL,
        #                             command=self.tree.xview, style='Dark.Horizontal.TScrollbar')
        # self.tree.configure(yscrollcommand=v_scrollbar.set,
        #                     xscrollcommand=h_scrollbar.set)

        self.tree.configure(yscrollcommand=v_scrollbar.set)
        
        # Grid treeview and scrollbars
        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        v_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        # h_scrollbar.grid(row=1, column=0, sticky=(tk.W, tk.E))

        # Bind right-click event for context menu
        self.tree.bind("<Button-3>", self.show_context_menu)

        # Status bar
        self.status_var = tk.StringVar()
        self.status_var.set("Ready")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var,
                               relief=tk.SUNKEN, anchor=tk.W,
                               style='Dark.TLabel')
        status_bar.grid(row=2, column=0, columnspan=3,
                        sticky=(tk.W, tk.E), pady=(10, 0))

    def import_data(self):
        """Import data from file"""
        file_types = [
            ('CSV files', '*.csv'),
        ]

        file_path = filedialog.askopenfilename(
            title="Select file to import",
            filetypes=file_types
        )

        if not file_path:
            return

        try:
            if file_path.endswith('.csv'):
                self.import_csv(file_path)
            else:
                messagebox.showerror("Error", "Unsupported file format")
                return

            self.refresh_treeview()
            self.status_var.set(f"Imported data from {
                                os.path.basename(file_path)}")

        except Exception as e:
            messagebox.showerror(
                "Import Error", f"Failed to import file:\n{str(e)}")
            self.status_var.set("Import failed")

    def import_csv(self, file_path):
        """Import data from CSV file"""
        # Reset auto-increment ID counter for new import
        self.next_id = 1

        # Clear existing data and used values
        self.data.clear()
        self.used_categories.clear()
        self.used_display_groups.clear()
        self.used_levels.clear()
        self.used_stacks.clear()

        imported_count = 0
        skipped_count = 0

        with open(file_path, 'r', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            for row in reader:
                # Convert data types and add auto-incrementing ID
                processed_row = self.process_csv_row(row)
                if processed_row is not None:
                    self.data.append(processed_row)
                    imported_count += 1
                else:
                    skipped_count += 1
                self.next_id += 1

        # Apply name standardization after import
        if imported_count > 0:
            self.standardize_names_silent()
            
            # Pre-fill all auto-generated fields
            self.prefill_auto_fields()
            
            # Auto-adjust column widths to fit data
            self.auto_adjust_column_widths()

        # Update status with import results
        self.status_var.set(f"Imported {imported_count} relics, skipped {
                            skipped_count} rows")

    def process_csv_row(self, row):
        """Process a CSV row and convert data types according to relic data specifications"""
        processed = {}

        # Auto-incrementing ID
        processed['id'] = self.next_id

        # Extract gameIds from multiple CSV columns
        game_ids = []
        game_id_columns = ['ID', 'passiveSpEffectId_1',
                           'passiveSpEffectId_2', 'passiveSpEffectId_3']

        for col in game_id_columns:
            value = row.get(col, '').strip()
            if value and value > '0':  # Skip empty, zero or negative values
                try:
                    game_ids.append(int(value))
                except (ValueError, TypeError):
                    pass

        # Remove duplicates and sort
        game_ids = sorted(list(set(game_ids)))
        processed['gameIds'] = ', '.join(
            map(str, game_ids)) if game_ids else ''

        # Process name with filtering
        name = str(row.get('Name', '')).strip()
        if name.startswith('Character Relic: '):
            processed['name'] = name[17:]  # Remove "Character Relic: " prefix
        elif name.startswith('Relic: '):
            processed['name'] = name[7:]   # Remove "Relic: " prefix
        else:
            # Skip rows that don't start with expected prefixes
            return None

        # Boolean fields with specific logic
        processed['debuff'] = self.str_to_bool(row.get('isDebuff', ''))

        # deep: isNumericEffect, 0 = true, 1 = false
        is_numeric = row.get('isNumericEffect', '').strip()
        processed['deep'] = (is_numeric == '0')

        # Initialize stacks as user-editable field (leave blank initially)
        processed['stacks'] = ''

        # Determine nightfarer from allow columns
        allow_columns = {
            'allowWylder': 'Wylder',
            'allowGuardian': 'Guardian',
            'allowIroneye': 'Ironeye',
            'allowDuchess': 'Duchess',
            'allowRaider': 'Raider',
            'allowRevenant': 'Revenant',
            'allowRecluse': 'Recluse',
            'allowExecutor': 'Executor'
        }

        nightfarer_values = []
        for col, name in allow_columns.items():
            if self.str_to_bool(row.get(col, '')):
                nightfarer_values.append(name)

        # Set nightfarer only if exactly one is true
        if len(nightfarer_values) == 1:
            processed['nightfarer'] = nightfarer_values[0]
        else:
            processed['nightfarer'] = ''

        # Level Group ID from attachFilterParamId
        try:
            attach_filter_param = row.get('attachFilterParamId', '').strip()
            processed['level_group_id'] = int(attach_filter_param) if attach_filter_param else 0
        except (ValueError, TypeError):
            processed['level_group_id'] = 0
        
        # Initialize new level_group field (user-editable)
        processed['level_group'] = ''
        
        # User-editable fields (leave blank initially)
        processed['category'] = ''
        processed['display_group'] = ''
        processed['level'] = ''

        return processed

    def str_to_bool(self, value):
        """Convert string to boolean"""
        if isinstance(value, bool):
            return value
        if isinstance(value, str):
            return value.lower() in ('true', '1', 'yes', 'on', 't', 'y')
        return bool(value)

    def export_data(self):
        """Export data to file"""
        if not self.data:
            messagebox.showwarning("No Data", "No data to export")
            return

        file_types = [
            ('JSON files', '*.json')
        ]

        file_path = filedialog.asksaveasfilename(
            title="Save data as",
            filetypes=file_types,
            defaultextension='.json'
        )

        if not file_path:
            return

        try:
            if file_path.endswith('.json'):
                self.export_json(file_path)
            else:
                messagebox.showerror("Error", "Unsupported file format")
                return

            self.status_var.set(f"Exported data to {
                                os.path.basename(file_path)}")

        except Exception as e:
            messagebox.showerror(
                "Export Error", f"Failed to export file:\n{str(e)}")
            self.status_var.set("Export failed")

    def export_json(self, file_path):
        """Export data to JSON file, with transformations:
        - exclude empty string fields
        - exclude 'id' and 'level_group_id' entirely
        - convert 'gameIds' (comma-separated string) to 'ids' (array[int])
        - convert 'stacks' ('Yes'/'No' string) to boolean (true/false)
        """
        # Filter and transform fields for each item
        filtered_data = []
        for item in self.data:
            filtered_item = {}
            for key, value in item.items():
                # Skip empty strings
                if value == "":
                    continue

                # Exclude internal 'id' and 'level_group_id' from export
                if key in ['id', 'level_group_id']:
                    continue

                # Transform 'gameIds' -> 'ids' as array of integers
                if key == 'gameIds':
                    try:
                        ids_set = self.parse_game_ids(value) if isinstance(value, str) else set()
                        ids_list = sorted(ids_set)
                        if ids_list:
                            filtered_item['ids'] = ids_list
                    except Exception:
                        # If parsing fails, omit ids gracefully
                        pass
                    continue

                # Transform 'stacks' -> boolean
                if key == 'stacks':
                    if value == 'Yes':
                        filtered_item[key] = True
                    elif value == 'No':
                        filtered_item[key] = False
                    # Skip if blank (empty string) - already handled by empty string check above
                    continue

                # Keep other values as-is
                filtered_item[key] = value
            filtered_data.append(filtered_item)
        
        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump(filtered_data, f, indent=2, ensure_ascii=False)

    def clear_data(self):
        """Clear all data from the treeview"""
        if self.data:
            result = messagebox.askyesno("Confirm Clear",
                                         "Are you sure you want to clear all data?")
            if result:
                self.data.clear()
                self.refresh_treeview()
                self.status_var.set("Data cleared")
        else:
            messagebox.showinfo("No Data", "No data to clear")

    def save_project(self):
        """Save current project state to file"""
        if not self.data:
            messagebox.showwarning("No Data", "No data to save")
            return

        file_path = filedialog.asksaveasfilename(
            title="Save Project As",
            filetypes=[('Project files', '*.rproj'), ('All files', '*.*')],
            defaultextension='.rproj'
        )

        if not file_path:
            return

        try:
            project_data = {
                'version': '1.0',
                'data': self.data,
                'next_id': self.next_id,
                'used_categories': list(self.used_categories),
                'used_display_groups': list(self.used_display_groups),
                'used_level_groups': list(self.used_level_groups),
                'used_levels': list(self.used_levels),
                'used_stacks': list(self.used_stacks),
                'sort_column': self.sort_column,
                'sort_reverse': self.sort_reverse
            }

            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(project_data, f, indent=2, ensure_ascii=False)

            self.status_var.set(f"Project saved to {
                                os.path.basename(file_path)}")

        except Exception as e:
            messagebox.showerror(
                "Save Error", f"Failed to save project:\n{str(e)}")
            self.status_var.set("Save failed")

    def migrate_project_data(self, project_data):
        """Migrate project data from older versions for backwards compatibility
        
        Handles the following field name changes:
        - stack_id -> level_group_id
        - stack_group -> level_group
        - used_stack_groups -> used_level_groups
        
        Returns tuple: (migrated_data, migration_performed)
        """
        # Check if this is an older version by looking for old field names
        needs_migration = False
        
        # Check data items for old field names
        if 'data' in project_data:
            for item in project_data['data']:
                if 'stack_id' in item or 'stack_group' in item:
                    needs_migration = True
                    break
        
        # Check used values for old field names
        if 'used_stack_groups' in project_data:
            needs_migration = True
        
        if not needs_migration:
            return project_data, False
        
        # Perform migration
        migrated_data = project_data.copy()
        
        # Migrate data items
        if 'data' in migrated_data:
            for item in migrated_data['data']:
                # Migrate stack_id -> level_group_id
                if 'stack_id' in item:
                    item['level_group_id'] = item.pop('stack_id')
                
                # Migrate stack_group -> level_group
                if 'stack_group' in item:
                    item['level_group'] = item.pop('stack_group')
        
        # Migrate used values
        if 'used_stack_groups' in migrated_data:
            migrated_data['used_level_groups'] = migrated_data.pop('used_stack_groups')
        
        return migrated_data, True

    def load_project(self):
        """Load project state from file"""
        file_path = filedialog.askopenfilename(
            title="Load Project",
            filetypes=[('Project files', '*.rproj'), ('All files', '*.*')]
        )

        if not file_path:
            return

        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                project_data = json.load(f)

            # Migrate project data for backwards compatibility
            project_data, migration_performed = self.migrate_project_data(project_data)

            # Validate project file format
            if 'data' not in project_data or 'next_id' not in project_data:
                messagebox.showerror(
                    "Invalid Project", "This is not a valid project file")
                return

            # Load data
            self.data = project_data['data']
            self.next_id = project_data.get('next_id', 1)

            # Load used values
            self.used_categories = set(project_data.get('used_categories', []))
            self.used_display_groups = set(project_data.get('used_display_groups', []))
            self.used_level_groups = set(project_data.get('used_level_groups', []))
            self.used_levels = set(project_data.get('used_levels', []))
            self.used_stacks = set(project_data.get('used_stacks', []))

            # Load sort state
            self.sort_column = project_data.get('sort_column')
            self.sort_reverse = project_data.get('sort_reverse', False)

            # Refresh display
            self.refresh_treeview()
            self.update_sort_indicators()
            
            # Auto-adjust column widths
            self.auto_adjust_column_widths()

            # Set status message
            if migration_performed:
                self.status_var.set(f"Project loaded from {
                                    os.path.basename(file_path)} (migrated from older version)")
            else:
                self.status_var.set(f"Project loaded from {
                                    os.path.basename(file_path)}")

        except Exception as e:
            messagebox.showerror(
                "Load Error", f"Failed to load project:\n{str(e)}")
            self.status_var.set("Load failed")

    def standardize_name(self, name, nightfarer=None):
        """Standardize a single name, handling all formatting patterns"""
        standardized_name = name

        # List of all possible nightfarer names to check for
        nightfarer_names = ['Wylder', 'Guardian', 'Ironeye',
                            'Duchess', 'Raider', 'Revenant', 'Recluse', 'Executor']

        # If nightfarer is provided, use it; otherwise check all possible nightfarers
        nightfarers_to_check = [nightfarer] if nightfarer else nightfarer_names

        # Check for various formatting patterns
        for nf in nightfarers_to_check:
            if nf and standardized_name.startswith(f"[{nf}] "):
                # Already properly formatted
                break
            elif standardized_name.startswith(f"{nf}: "):
                # Format: "Nightfarer: name" -> "[Nightfarer] name"
                standardized_name = f"[{nf}] {standardized_name[len(nf) + 2:]}"
                break
            elif standardized_name.startswith(f"{nf} "):
                # Format: "Nightfarer name" -> "[Nightfarer] name"
                standardized_name = f"[{nf}] {standardized_name[len(nf) + 1:]}"
                break

        # If nightfarer is provided and no pattern was matched, add the nightfarer prefix
        if nightfarer and not any(standardized_name.startswith(f"[{nf}] ") for nf in nightfarers_to_check):
            standardized_name = f"[{nightfarer}] {standardized_name}"

        # Replace "- " with ", " (but not "-" without space)
        standardized_name = standardized_name.replace("- ", ", ")

        return standardized_name

    def standardize_names_silent(self):
        """Standardize names and merge duplicates without showing status messages (for import)"""
        # First pass: Apply nightfarer formatting to rows with nightfarer values
        for item in self.data:
            nightfarer = item.get('nightfarer', '').strip()
            name = item.get('name', '').strip()

            if nightfarer and name:
                # Apply nightfarer formatting and dash replacement
                standardized_name = self.standardize_name(name, nightfarer)
                item['name'] = standardized_name

        # Second pass: Check for edge cases in names without nightfarer patterns
        for item in self.data:
            name = item.get('name', '').strip()

            if name:
                # Check if name already has nightfarer formatting
                has_nightfarer_format = any(name.startswith(f"[{nf}] ") for nf in [
                                            'Wylder', 'Guardian', 'Ironeye', 'Duchess', 'Raider', 'Revenant', 'Recluse', 'Executor'])

                if not has_nightfarer_format:
                    # Apply edge case standardization (nightfarer patterns in name)
                    standardized_name = self.standardize_name(name)
                    item['name'] = standardized_name

        # Third pass: detect and merge duplicates
        self.merge_duplicate_names()

        # Refresh display
        self.refresh_treeview()

    def prefill_auto_fields(self):
        """Pre-fill all auto-generated fields: nightfarer display groups, debuff categories, stack groups, and levels"""
        # First pass: Pre-fill nightfarer display groups and debuff categories
        for item in self.data:
            name = item.get('name', '').strip()
            if name.startswith('[') and ']' in name:
                # Extract nightfarer name from [Nightfarer] pattern
                end_bracket = name.find(']')
                if end_bracket > 1:
                    nightfarer_name = name[1:end_bracket].strip()
                    if nightfarer_name:
                        item['display_group'] = nightfarer_name
            
            # Pre-fill category for debuff effects
            if item.get('debuff', False) and not item.get('category', '').strip():
                item['category'] = 'Debuff'
        
        # Second pass: Group items by level_group_id for common text pre-filling
        items_by_level_group_id = {}
        for item in self.data:
            level_group_id = item.get('level_group_id', 0)
            if level_group_id not in items_by_level_group_id:
                items_by_level_group_id[level_group_id] = []
            items_by_level_group_id[level_group_id].append(item)
        
        # Process each level group
        for level_group_id, items in items_by_level_group_id.items():
            if len(items) > 1:  # Only process groups with multiple items
                # Find common text for all items in this stack group
                common_text = self.find_common_text([item.get('name', '') for item in items])
                if common_text:
                    # Apply common text to level_group field for all items
                    for item in items:
                        item['level_group'] = common_text
                
                # Pre-fill levels based on "+" patterns
                self.prefill_levels_for_level_group(items)

    def prefill_levels_for_level_group(self, items):
        """Pre-fill levels based on '+' patterns in names for a level group"""
        if len(items) < 2:
            return
        
        # Extract level information from names
        level_info = []
        for item in items:
            name = item.get('name', '').strip()
            level = self.extract_level_from_name(name)
            level_info.append((item, level))
        
        # Check if we have mixed patterns (some with no number, some with numbers)
        has_no_number = any(level is None for _, level in level_info)
        has_numbers = any(level is not None for _, level in level_info)
        
        if has_no_number and has_numbers:
            # Mixed pattern: no number = level 1, +1 = level 2, +2 = level 3, etc.
            for item, level in level_info:
                if level is None:
                    item['level'] = 1
                else:
                    item['level'] = level + 1
        elif has_numbers and not has_no_number:
            # All have numbers: +1 = level 1, +2 = level 2, etc.
            for item, level in level_info:
                if level is not None:
                    item['level'] = level
        # If all have no numbers, leave levels as they are (0)

    def extract_level_from_name(self, name):
        """Extract level number from name ending with '+N' pattern"""
        import re
        # Look for pattern like " +1", " +2", etc. at the end
        match = re.search(r' \+(\d+)$', name)
        if match:
            return int(match.group(1))
        return None

    def find_common_text(self, names):
        """Find common text among a list of names"""
        if not names or len(names) < 2:
            return ""
        
        # Remove nightfarer prefixes for comparison
        def remove_nightfarer_prefix(name):
            for nf in ['Wylder', 'Guardian', 'Ironeye', 'Duchess', 'Raider', 'Revenant', 'Recluse', 'Executor']:
                if name.startswith(f"[{nf}] "):
                    return name[len(nf) + 3:]
            return name
        
        # Clean names and remove empty ones
        clean_names = [remove_nightfarer_prefix(name).strip() for name in names if name.strip()]
        if len(clean_names) < 2:
            return ""
        
        # Find common prefix
        common_prefix = self.find_common_prefix(clean_names)
        
        # Find common suffix
        common_suffix = self.find_common_suffix(clean_names)
        
        # Find common words
        common_words = self.find_common_words(clean_names)
        
        # Find longest common substring
        common_substring = self.find_longest_common_substring(clean_names)
        
        # Return the longest common text found
        candidates = [common_prefix, common_suffix, common_words, common_substring]
        candidates = [c for c in candidates if c and len(c.strip()) > 0]
        
        if candidates:
            # Return the longest common text
            result = max(candidates, key=len).strip()
            
            # Ensure first character is uppercase
            if result and result[0].islower():
                result = result[0].upper() + result[1:]
            
            return result
        
        return ""

    def find_longest_common_substring(self, names):
        """Find the longest common substring among names (case-insensitive)"""
        if not names or len(names) < 2:
            return ""
        
        # Use the shortest name as the base for substring generation
        shortest = min(names, key=len)
        if len(shortest) < 3:  # Need at least 3 characters for a meaningful substring
            return ""
        
        longest_common = ""
        longest_common_original_case = ""
        
        # Try all possible substrings of the shortest name
        for i in range(len(shortest)):
            for j in range(i + 3, len(shortest) + 1):  # Minimum 3 characters
                substring = shortest[i:j]
                
                # Check if this substring appears in all names (case-insensitive)
                if all(substring.lower() in name.lower() for name in names):
                    # Only consider substrings that end at word boundaries or are substantial
                    if (len(substring) > len(longest_common) and 
                        (substring.endswith(' ') or substring.endswith('-') or 
                         substring.endswith(',') or len(substring) >= 10)):
                        longest_common = substring
                        # Find the most common case version across all names
                        case_versions = []
                        for name in names:
                            if substring.lower() in name.lower():
                                # Find the position and extract with original case
                                pos = name.lower().find(substring.lower())
                                if pos != -1:
                                    case_versions.append(name[pos:pos+len(substring)])
                        
                        # Use the most common case version, or the first one if tied
                        if case_versions:
                            # Count occurrences of each case version
                            from collections import Counter
                            case_counts = Counter(case_versions)
                            longest_common_original_case = case_counts.most_common(1)[0][0]
        
        # Clean up the result
        if longest_common_original_case:
            # Remove trailing "+ " if present
            if longest_common_original_case.rstrip().endswith('+'):
                longest_common_original_case = longest_common_original_case.rstrip().rstrip('+').rstrip()
            
            # Ensure first character is uppercase (only if we have content after trimming)
            if longest_common_original_case and longest_common_original_case.strip() and longest_common_original_case.strip()[0].islower():
                longest_common_original_case = longest_common_original_case.strip()
                longest_common_original_case = longest_common_original_case[0].upper() + longest_common_original_case[1:]
            
            return longest_common_original_case.rstrip()
        
        return ""

    def find_common_prefix(self, names):
        """Find common prefix among names"""
        if not names:
            return ""
        
        # Start with the shortest name
        shortest = min(names, key=len)
        common_prefix = ""
        
        for i in range(len(shortest)):
            char = shortest[i]
            if all(name[i] == char for name in names if i < len(name)):
                common_prefix += char
            else:
                break
        
        # Only return if it ends at a word boundary or is substantial
        if len(common_prefix) > 3 and (common_prefix.endswith(' ') or common_prefix.endswith('-') or common_prefix.endswith(',')):
            return common_prefix.rstrip()
        
        return ""

    def find_common_suffix(self, names):
        """Find common suffix among names"""
        if not names:
            return ""
        
        # Start with the shortest name
        shortest = min(names, key=len)
        common_suffix = ""
        
        for i in range(1, len(shortest) + 1):
            char = shortest[-i]
            if all(name[-i] == char for name in names if i <= len(name)):
                common_suffix = char + common_suffix
            else:
                break
        
        # Only return if it starts at a word boundary or is substantial
        if len(common_suffix) > 3 and (common_suffix.startswith(' ') or common_suffix.startswith('-') or common_suffix.startswith(',')):
            return common_suffix.lstrip()
        
        return ""

    def find_common_words(self, names):
        """Find common words among names"""
        if not names:
            return ""
        
        # Split each name into words
        word_lists = [name.split() for name in names]
        
        if not word_lists or not word_lists[0]:
            return ""
        
        # Find common words at the beginning
        common_words = []
        min_words = min(len(words) for words in word_lists)
        
        for i in range(min_words):
            word = word_lists[0][i]
            if all(words[i] == word for words in word_lists):
                common_words.append(word)
            else:
                break
        
        if common_words:
            return ' '.join(common_words)
        
        return ""

    def merge_duplicate_names(self):
        """Merge duplicate names and truncated name variants by keeping lowest ID and combining gameIDs"""
        # First pass: Group by exact name matches
        name_groups = {}
        for item in self.data:
            name = item.get('name', '').strip()
            if name:
                if name not in name_groups:
                    name_groups[name] = []
                name_groups[name].append(item)

        # Process exact duplicates first
        items_to_remove = []
        for name, items in name_groups.items():
            if len(items) > 1:
                self.merge_items(items, items_to_remove)

        # Second pass: Check for truncated name variants
        remaining_items = [item for item in self.data if item not in items_to_remove]
        self.merge_truncated_names(remaining_items, items_to_remove)

        # Remove duplicate items
        for item in items_to_remove:
            if item in self.data:
                self.data.remove(item)

    def merge_items(self, items, items_to_remove):
        """Merge a group of items, keeping the one with lowest ID and combining gameIDs"""
        # Sort by ID to get the one with lowest ID
        items.sort(key=lambda x: x.get('id', 0))
        keep_item = items[0]  # Keep the one with lowest ID

        # Collect all unique gameIDs
        all_game_ids = set()
        for item in items:
            game_ids_str = item.get('gameIds', '').strip()
            if game_ids_str:
                # Parse comma-separated game IDs
                try:
                    game_ids = [int(x.strip())
                                for x in game_ids_str.split(',') if x.strip()]
                    all_game_ids.update(game_ids)
                except (ValueError, TypeError):
                    pass

        # Update the kept item with combined gameIDs
        if all_game_ids:
            keep_item['gameIds'] = ', '.join(map(str, sorted(all_game_ids)))

        # Mark other items for removal
        items_to_remove.extend(items[1:])

    def merge_truncated_names(self, items, items_to_remove):
        """Check for truncated name variants and merge them"""
        # Group items by their first few words to find potential matches
        potential_groups = {}
        
        for item in items:
            name = item.get('name', '').strip()
            if name:
                # Remove nightfarer prefix for grouping
                clean_name = self.remove_nightfarer_prefix_for_grouping(name)
                words = clean_name.split()
                if len(words) >= 2:
                    key = ' '.join(words[:2])  # Use first 2 words as grouping key
                    if key not in potential_groups:
                        potential_groups[key] = []
                    potential_groups[key].append(item)

        # Check each group for truncated variants
        for key, group_items in potential_groups.items():
            if len(group_items) >= 2:
                # Check all pairs for truncation relationships
                for i in range(len(group_items)):
                    for j in range(i + 1, len(group_items)):
                        item1 = group_items[i]
                        item2 = group_items[j]
                        
                        if item1 not in items_to_remove and item2 not in items_to_remove:
                            name1 = item1.get('name', '').strip()
                            name2 = item2.get('name', '').strip()
                            
                            # Check if one is a prefix of the other
                            if self.is_truncated_variant(name1, name2):
                                # Check if they share any gameIDs
                                if self.share_game_ids(item1, item2):
                                    # Merge them
                                    self.merge_truncated_pair(item1, item2, items_to_remove)

    def remove_nightfarer_prefix_for_grouping(self, name):
        """Remove nightfarer prefix for grouping purposes"""
        for nf in ['Wylder', 'Guardian', 'Ironeye', 'Duchess', 'Raider', 'Revenant', 'Recluse', 'Executor']:
            if name.startswith(f"[{nf}] "):
                return name[len(nf) + 3:]
        return name

    def is_truncated_variant(self, name1, name2):
        """Check if one name is a truncated version of the other"""
        # Remove nightfarer prefix for comparison
        def remove_nightfarer_prefix(name):
            for nf in ['Wylder', 'Guardian', 'Ironeye', 'Duchess', 'Raider', 'Revenant', 'Recluse', 'Executor']:
                if name.startswith(f"[{nf}] "):
                    return name[len(nf) + 3:]
            return name
        
        clean_name1 = remove_nightfarer_prefix(name1)
        clean_name2 = remove_nightfarer_prefix(name2)
        
        # First check: if names are identical after removing nightfarer prefixes
        if clean_name1 == clean_name2:
            return True
        
        # Check if one is a prefix of the other (with word boundary)
        if clean_name1.startswith(clean_name2 + " ") or clean_name2.startswith(clean_name1 + " "):
            return True
        
        # Check if one is significantly shorter and contained in the other
        if len(clean_name1) > len(clean_name2) * 1.5 and clean_name1.startswith(clean_name2):
            return True
        if len(clean_name2) > len(clean_name1) * 1.5 and clean_name2.startswith(clean_name1):
            return True
        
        # Check for word-level differences (like "permanently" vs missing word)
        if self.are_word_variants(clean_name1, clean_name2):
            return True
            
        return False

    def are_word_variants(self, name1, name2):
        """Check if two names are variants with word differences"""
        words1 = name1.split()
        words2 = name2.split()
        
        # If one has significantly more words, check if it contains the other
        if len(words1) > len(words2) + 1:
            # Check if words2 is a subset of words1 in order
            return self.is_word_subset(words2, words1)
        elif len(words2) > len(words1) + 1:
            # Check if words1 is a subset of words2 in order
            return self.is_word_subset(words1, words2)
        
        # Check for single word differences
        if abs(len(words1) - len(words2)) == 1:
            shorter = words1 if len(words1) < len(words2) else words2
            longer = words2 if len(words1) < len(words2) else words1
            
            # Check if shorter is contained in longer with at most one word difference
            return self.has_single_word_difference(shorter, longer)
        
        return False

    def is_word_subset(self, shorter_words, longer_words):
        """Check if shorter word list is a subset of longer word list in order"""
        if len(shorter_words) == 0:
            return True
        
        i = 0
        for word in longer_words:
            if i < len(shorter_words) and word == shorter_words[i]:
                i += 1
                if i == len(shorter_words):
                    return True
        
        return i == len(shorter_words)

    def has_single_word_difference(self, shorter, longer):
        """Check if longer has at most one extra word compared to shorter"""
        if len(longer) != len(shorter) + 1:
            return False
        
        # Check if shorter is contained in longer with one word inserted
        i = j = 0
        extra_words = 0
        
        while i < len(shorter) and j < len(longer):
            if shorter[i] == longer[j]:
                i += 1
                j += 1
            else:
                j += 1
                extra_words += 1
                if extra_words > 1:
                    return False
        
        return i == len(shorter) and extra_words <= 1

    def share_game_ids(self, item1, item2):
        """Check if two items share any gameIDs"""
        game_ids1 = self.parse_game_ids(item1.get('gameIds', ''))
        game_ids2 = self.parse_game_ids(item2.get('gameIds', ''))
        
        return bool(game_ids1.intersection(game_ids2))

    def parse_game_ids(self, game_ids_str):
        """Parse game IDs string into a set of integers"""
        game_ids = set()
        if game_ids_str:
            try:
                game_ids = set(int(x.strip()) for x in game_ids_str.split(',') if x.strip())
            except (ValueError, TypeError):
                pass
        return game_ids

    def merge_truncated_pair(self, item1, item2, items_to_remove):
        """Merge two items that are truncated variants"""
        # Keep the one with lower ID
        if item1.get('id', 0) < item2.get('id', 0):
            keep_item = item1
            remove_item = item2
        else:
            keep_item = item2
            remove_item = item1
        
        # Use the longer name
        name1 = item1.get('name', '').strip()
        name2 = item2.get('name', '').strip()
        if len(name2) > len(name1):
            keep_item['name'] = name2
        
        # Combine gameIDs
        game_ids1 = self.parse_game_ids(item1.get('gameIds', ''))
        game_ids2 = self.parse_game_ids(item2.get('gameIds', ''))
        all_game_ids = game_ids1.union(game_ids2)
        
        if all_game_ids:
            keep_item['gameIds'] = ', '.join(map(str, sorted(all_game_ids)))
        
        # Mark the other item for removal
        items_to_remove.append(remove_item)

    def refresh_treeview(self):
        """Refresh the treeview with current data"""
        # Clear existing items
        for item in self.tree.get_children():
            self.tree.delete(item)

        # Add data to treeview
        for i, item in enumerate(self.data):
            values = (
                item.get('id', ''),
                item.get('gameIds', ''),
                item.get('name', ''),
                item.get('category', ''),
                item.get('display_group', ''),
                item.get('level_group_id', 0),
                item.get('level_group', ''),
                item.get('level', ''),
                item.get('nightfarer', ''),
                'Yes' if item.get('deep', False) else 'No',
                'Yes' if item.get('debuff', False) else 'No',
                item.get('stacks', '')
            )
            self.tree.insert('', 'end', values=values, tags=(
                'even' if i % 2 == 0 else 'odd',))

        # Configure alternating row colors for dark theme
        self.tree.tag_configure('even', background='#3c3c3c')
        self.tree.tag_configure('odd', background='#404040')
        self.tree.tag_configure('selected', background='#0078d4')

    def auto_adjust_column_widths(self):
        """Automatically adjust column widths based on content"""
        if not self.data:
            return
        
        # Define columns and their minimum widths
        columns = ['id', 'gameIds', 'name', 'category', 'display_group',
                   'level_group_id', 'level_group', 'level', 'nightfarer', 'deep', 'debuff', 'stacks']
        
        min_widths = {
            'id': 30,
            'gameIds': 80,
            'name': 120,
            'category': 80,
            'display_group': 80,
            'level_group_id': 80,
            'level_group': 80,
            'level': 50,
            'nightfarer': 80,
            'deep': 50,
            'debuff': 50,
            'stacks': 50
        }
        
        # Calculate maximum width needed for each column
        for column in columns:
            max_width = min_widths[column]
            
            # Check header width
            header_width = len(self.tree.heading(column)['text']) * 8  # Approximate character width
            max_width = max(max_width, header_width)
            
            # Check data widths
            for item in self.data:
                value = item.get(column, '')
                
                # Convert value to string and handle special cases
                if column in ['deep', 'debuff']:
                    display_value = 'Yes' if value else 'No'
                else:
                    display_value = str(value)
                
                # Calculate width needed for this value
                text_width = len(display_value) * 8  # Approximate character width
                max_width = max(max_width, text_width)
            
            # Set column width with some padding
            final_width = min(max_width + 20, 600)  # Add padding, cap at 600px
            self.tree.column(column, width=final_width)

    def sort_by_column(self, column):
        """Sort data by the specified column"""
        # Toggle sort direction if clicking the same column
        if self.sort_column == column:
            self.sort_reverse = not self.sort_reverse
        else:
            self.sort_reverse = False
            self.sort_column = column

        # Sort the data
        self.data.sort(key=lambda x: self.get_sort_key(
            x, column), reverse=self.sort_reverse)

        # Update the treeview
        self.refresh_treeview()

        # Update column header to show sort direction
        self.update_sort_indicators()

    def get_sort_key(self, item, column):
        """Get the sort key for an item based on column"""
        value = item.get(column, '')

        # Handle different data types
        if column == 'id':
            return int(value) if value else 0
        elif column in ['debuff', 'deep', 'stacks']:
            # Boolean columns - sort True before False
            return bool(value)
        elif column in ['level', 'level_group_id']:
            # Integer columns - handle empty strings
            try:
                return int(value) if value else 0
            except (ValueError, TypeError):
                return 0
        elif column == 'level_group':
            # String column - case insensitive
            return str(value).lower()
        elif column == 'gameIds':
            # Game IDs - sort by first ID in the list
            if value:
                try:
                    first_id = value.split(',')[0].strip()
                    return int(first_id) if first_id else 0
                except (ValueError, TypeError):
                    return 0
            return 0
        else:
            # String columns - case insensitive
            return str(value).lower()

    def update_sort_indicators(self):
        """Update column headers to show sort direction"""
        # Reset all headers
        for col in ['id', 'gameIds', 'name', 'debuff', 'deep', 'stacks', 'level_group', 'nightfarer', 'category', 'display_group', 'level']:
            if col == 'id':
                text = 'ID'
            elif col == 'gameIds':
                text = 'Game IDs'
            elif col == 'name':
                text = 'Name'
            elif col == 'debuff':
                text = 'Debuff'
            elif col == 'deep':
                text = 'Deep'
            elif col == 'stacks':
                text = 'Stacks'
            elif col == 'level_group':
                text = 'Level Group'
            elif col == 'nightfarer':
                text = 'Nightfarer'
            elif col == 'category':
                text = 'Category'
            elif col == 'display_group':
                text = 'Display Group'
            elif col == 'level':
                text = 'Level'

            # Add sort indicator
            if col == self.sort_column:
                indicator = ' ↓' if self.sort_reverse else ' ↑'
                text += indicator

            self.tree.heading(col, text=text)

    def show_context_menu(self, event):
        """Show context menu for editable columns"""
        # Get the item and column under the cursor
        item = self.tree.identify_row(event.y)
        column = self.tree.identify_column(event.x)

        if not item or not column:
            return

        # Get column name from column index
        column_index = int(column.replace('#', '')) - 1
        columns = ['id', 'gameIds', 'name', 'category', 'display_group',
                   'level_group_id', 'level_group', 'level', 'nightfarer', 'deep', 'debuff', 'stacks']

        if column_index < 0 or column_index >= len(columns):
            return

        column_name = columns[column_index]

        # Only show context menu for editable columns
        if column_name not in ['category', 'display_group', 'level_group', 'level', 'stacks']:
            return

        # Get selected items (including the clicked item if not selected)
        selected_items = list(self.tree.selection())
        if item not in selected_items:
            selected_items = [item]

        # Create context menu
        context_menu = tk.Menu(self.root, tearoff=0, bg='#404040', fg='#ffffff',
                               activebackground='#505050', activeforeground='#ffffff')

        # Add "New [column]" option with count if multiple selected
        display_name = column_name.replace('_', ' ').title()
        if len(selected_items) > 1:
            context_menu.add_command(
                label=f"New {display_name} ({len(selected_items)} rows)",
                command=lambda: self.edit_cell_value(
                    selected_items, column_name, "")
            )
        else:
            context_menu.add_command(
                label=f"New {display_name}",
                command=lambda: self.edit_cell_value(
                    selected_items, column_name, "")
            )

        # Add "Clear" option
        if len(selected_items) > 1:
            context_menu.add_command(
                label=f"Clear {display_name} ({len(selected_items)} rows)",
                command=lambda: self.set_cell_value(
                    selected_items, column_name, "")
            )
        else:
            context_menu.add_command(
                label=f"Clear {display_name}",
                command=lambda: self.set_cell_value(
                    selected_items, column_name, "")
            )

        # Add special option for display_group column
        if column_name == 'display_group':
            # Check if all selected rows have the same level_group
            level_groups = set()
            for item in selected_items:
                item_index = self.tree.index(item)
                if 0 <= item_index < len(self.data):
                    level_group = self.data[item_index].get('level_group', '').strip()
                    if level_group:
                        level_groups.add(level_group)
            
            if len(level_groups) == 1 and level_groups:
                # All selected rows have the same non-empty level_group
                level_group_value = list(level_groups)[0]
                if len(selected_items) > 1:
                    context_menu.add_command(
                        label=f"Use Level Group: {level_group_value} ({len(selected_items)} rows)",
                        command=lambda: self.set_cell_value(
                            selected_items, column_name, level_group_value)
                    )
                else:
                    context_menu.add_command(
                        label=f"Use Level Group: {level_group_value}",
                        command=lambda: self.set_cell_value(
                            selected_items, column_name, level_group_value)
                    )
                context_menu.add_separator()

        # Add special options for stacks column
        if column_name == 'stacks':
            context_menu.add_separator()
            # Add Yes option
            if len(selected_items) > 1:
                context_menu.add_command(
                    label=f"Set to Yes ({len(selected_items)} rows)",
                    command=lambda: self.set_cell_value(
                        selected_items, column_name, "Yes")
                )
            else:
                context_menu.add_command(
                    label="Set to Yes",
                    command=lambda: self.set_cell_value(
                        selected_items, column_name, "Yes")
                )
            
            # Add No option
            if len(selected_items) > 1:
                context_menu.add_command(
                    label=f"Set to No ({len(selected_items)} rows)",
                    command=lambda: self.set_cell_value(
                        selected_items, column_name, "No")
                )
            else:
                context_menu.add_command(
                    label="Set to No",
                    command=lambda: self.set_cell_value(
                        selected_items, column_name, "No")
                )
            
            context_menu.add_separator()

        # Add separator
        context_menu.add_separator()

        # Add previously used values
        used_values = self.get_used_values(column_name)
        if used_values:
            for value in sorted(used_values):
                if len(selected_items) > 1:
                    context_menu.add_command(
                        label=f"{value} ({len(selected_items)} rows)",
                        command=lambda v=value: self.set_cell_value(
                            selected_items, column_name, v)
                    )
                else:
                    context_menu.add_command(
                        label=value,
                        command=lambda v=value: self.set_cell_value(
                            selected_items, column_name, v)
                    )
        else:
            context_menu.add_command(
                label="No previous values", state="disabled")

        # Show context menu
        try:
            context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            context_menu.grab_release()

    def get_used_values(self, column_name):
        """Get previously used values for a column"""
        if column_name == 'category':
            return self.used_categories
        elif column_name == 'display_group':
            return self.used_display_groups
        elif column_name == 'level_group':
            return self.used_level_groups
        elif column_name == 'level':
            return self.used_levels
        elif column_name == 'stacks':
            return self.used_stacks
        return set()

    def edit_cell_value(self, items, column_name, current_value):
        """Open dialog to edit cell value for multiple items"""
        # Create dialog window
        dialog = tk.Toplevel(self.root)
        display_name = column_name.replace('_', ' ').title()
        if len(items) > 1:
            dialog.title(f"Edit {display_name} ({len(items)} rows)")
        else:
            dialog.title(f"Edit {display_name}")
        dialog.configure(bg='#2b2b2b')
        dialog.geometry("300x150")
        dialog.resizable(False, False)

        # Center the dialog
        dialog.transient(self.root)
        dialog.grab_set()

        # Create input frame
        input_frame = ttk.Frame(dialog, style='Dark.TFrame')
        input_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # Label
        if len(items) > 1:
            label_text = f"Enter new {display_name} for {
                len(items)} selected rows:"
        else:
            label_text = f"Enter new {display_name}:"
        label = ttk.Label(input_frame, text=label_text, style='Dark.TLabel')
        label.pack(pady=(0, 10))

        # Entry field
        entry_var = tk.StringVar(value=current_value)
        entry = ttk.Entry(input_frame, textvariable=entry_var,
                          font=('Arial', 10))
        entry.pack(fill=tk.X, pady=(0, 20))
        entry.focus()
        entry.select_range(0, tk.END)

        # Button frame
        button_frame = ttk.Frame(input_frame, style='Dark.TFrame')
        button_frame.pack(fill=tk.X)

        # Buttons
        ttk.Button(button_frame, text="Cancel", command=dialog.destroy,
                   style='Modern.TButton').pack(side=tk.RIGHT, padx=(5, 0))
        ttk.Button(button_frame, text="OK", command=lambda: self.save_cell_value(
            dialog, items, column_name, entry_var.get()), style='Modern.TButton').pack(side=tk.RIGHT)

        # Bind Enter key to OK button
        entry.bind('<Return>', lambda e: self.save_cell_value(
            dialog, items, column_name, entry_var.get()))

    def save_cell_value(self, dialog, items, column_name, new_value):
        """Save the new cell value for multiple items"""
        if not new_value.strip():
            dialog.destroy()
            return

        updated_count = 0

        # Update all selected items
        for item in items:
            item_index = self.tree.index(item)
            if 0 <= item_index < len(self.data):
                self.data[item_index][column_name] = new_value.strip()
                updated_count += 1

        # Add to used values
        if column_name == 'category':
            self.used_categories.add(new_value.strip())
        elif column_name == 'display_group':
            self.used_display_groups.add(new_value.strip())
        elif column_name == 'level_group':
            self.used_level_groups.add(new_value.strip())
        elif column_name == 'level':
            self.used_levels.add(new_value.strip())
        elif column_name == 'stacks':
            self.used_stacks.add(new_value.strip())

        # Refresh the treeview
        self.refresh_treeview()
        
        # Auto-adjust column widths after data change
        self.auto_adjust_column_widths()

        # Update status
        if updated_count > 1:
            self.status_var.set(f"Updated {column_name} to '{
                                new_value.strip()}' for {updated_count} rows")
        else:
            self.status_var.set(f"Updated {column_name} to '{
                                new_value.strip()}'")

        dialog.destroy()

    def set_cell_value(self, items, column_name, value):
        """Set cell value directly from context menu for multiple items"""
        updated_count = 0

        # Update all selected items
        for item in items:
            item_index = self.tree.index(item)
            if 0 <= item_index < len(self.data):
                self.data[item_index][column_name] = value
                updated_count += 1

        # Refresh the treeview
        self.refresh_treeview()
        
        # Auto-adjust column widths after data change
        self.auto_adjust_column_widths()

        # Update status
        if updated_count > 1:
            self.status_var.set(f"Set {column_name} to '{
                                value}' for {updated_count} rows")
        else:
            self.status_var.set(f"Set {column_name} to '{value}'")


def main():
    """Main function to run the application"""
    root = tk.Tk()
    # Maximize window based on OS
    try:
        import platform
        system_name = platform.system().lower()
        if 'windows' in system_name:
            # Windows
            root.state('zoomed')
        else:
            # Linux (and others treated as Linux per requirement)
            root.attributes('-zoomed', True)
    except Exception:
        # Fallback to Linux behavior if detection fails
        root.attributes('-zoomed', True)
    app = DataManagerGUI(root)

    root.mainloop()


if __name__ == "__main__":
    main()
