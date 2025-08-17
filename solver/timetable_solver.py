if vars_at_time:
                    self.model.Add(sum(vars_at_time) <= 1)
                    
    def _add_teacher_availability(self, availability_data):
        """Add teacher availability constraints."""
        for teacher_id, unavailable_slots in availability_data.items():
            for slot_id in unavailable_slots:
                for key, var in self.assignments.items():
                    if key[0] == teacher_id and key[3] == slot_id:
                        self.model.Add(var == 0)
                        
    def _add_room_capacity(self):
        """Ensure room capacity is not exceeded."""
        rooms = self.data.get('rooms', {})
        classes = self.data.get('classes', {})
        
        for key, var in self.assignments.items():
            _, class_id, room_id, _, _ = key
            if class_id in classes and room_id in rooms:
                class_size = classes[class_id].get('size', 30)
                room_capacity = rooms[room_id].get('capacity', 30)
                if class_size > room_capacity:
                    self.model.Add(var == 0)
                    
    def _add_minimize_gaps(self, weight):
        """Minimize gaps in teacher schedules (soft constraint)."""
        # Implementation for gap minimization
        pass
        
    def _add_preferred_times(self, preferences, weight):
        """Add preferred time slots for teachers (soft constraint)."""
        # Implementation for preferred times
        pass
        
    def _configure_solver(self):
        """Configure solver parameters."""
        self.solver.parameters.max_time_in_seconds = self.time_limit
        self.solver.parameters.num_search_workers = self.num_workers
        
        if self.search_strategy == "First Unbound Minimum Domain":
            self.solver.parameters.search_branching = cp_model.FIRST_UNBOUND_MIN_VALUE
        elif self.search_strategy == "Random":
            self.solver.parameters.random_seed = 42
            
    def _extract_solution(self) -> Dict:
        """Extract solution from solved model."""
        solution = {
            'assignments': [],
            'statistics': {
                'total_assignments': 0,
                'conflicts': 0,
                'utilization': 0.0,
                'solve_time': self.solver.WallTime()
            }
        }
        
        for key, var in self.assignments.items():
            if self.solver.Value(var) == 1:
                teacher_id, class_id, room_id, slot_id, subject_id = key
                solution['assignments'].append({
                    'teacher': teacher_id,
                    'class': class_id,
                    'room': room_id,
                    'timeslot': slot_id,
                    'subject': subject_id
                })
                solution['statistics']['total_assignments'] += 1
                
        return solution


# ============================================
# importer/excel_importer.py
# ============================================
"""
Excel importer for legacy timetable data
"""

import pandas as pd
from typing import Dict, Any
from pathlib import Path


class ExcelImporter:
    """Import timetable data from legacy Excel files."""
    
    def import_file(self, file_path: str) -> Dict[str, Any]:
        """Import data from Excel file with multiple sheets."""
        data = {}
        
        # Read all sheets
        excel_file = pd.ExcelFile(file_path)
        
        # Import teachers
        if 'Teachers' in excel_file.sheet_names:
            df = pd.read_excel(excel_file, 'Teachers')
            data['teachers'] = self._process_teachers(df)
            
        # Import classes
        if 'Classes' in excel_file.sheet_names:
            df = pd.read_excel(excel_file, 'Classes')
            data['classes'] = self._process_classes(df)
            
        # Import rooms
        if 'Rooms' in excel_file.sheet_names:
            df = pd.read_excel(excel_file, 'Rooms')
            data['rooms'] = self._process_rooms(df)
            
        # Import subjects
        if 'Subjects' in excel_file.sheet_names:
            df = pd.read_excel(excel_file, 'Subjects')
            data['subjects'] = self._process_subjects(df)
            
        # Import timeslots
        if 'Timeslots' in excel_file.sheet_names:
            df = pd.read_excel(excel_file, 'Timeslots')
            data['timeslots'] = self._process_timeslots(df)
            
        return data
        
    def _process_teachers(self, df: pd.DataFrame) -> Dict:
        """Process teachers dataframe."""
        teachers = {}
        for _, row in df.iterrows():
            teacher_id = str(row.get('ID', row.get('Teacher ID', '')))
            if teacher_id:
                teachers[teacher_id] = {
                    'name': row.get('Name', ''),
                    'subjects': str(row.get('Subjects', '')).split(','),
                    'max_hours': int(row.get('Max Hours', 20)),
                    'availability': str(row.get('Availability', 'All'))
                }
        return teachers
        
    def _process_classes(self, df: pd.DataFrame) -> Dict:
        """Process classes dataframe."""
        classes = {}
        for _, row in df.iterrows():
            class_id = str(row.get('ID', row.get('Class ID', '')))
            if class_id:
                classes[class_id] = {
                    'name': row.get('Name', ''),
                    'size': int(row.get('Size', 30)),
                    'level': row.get('Level', ''),
                    'subjects': str(row.get('Subjects', '')).split(',')
                }
        return classes
        
    def _process_rooms(self, df: pd.DataFrame) -> Dict:
        """Process rooms dataframe."""
        rooms = {}
        for _, row in df.iterrows():
            room_id = str(row.get('ID', row.get('Room ID', '')))
            if room_id:
                rooms[room_id] = {
                    'name': row.get('Name', ''),
                    'capacity': int(row.get('Capacity', 30)),
                    'type': row.get('Type', 'Classroom'),
                    'facilities': str(row.get('Facilities', '')).split(',')
                }
        return rooms
        
    def _process_subjects(self, df: pd.DataFrame) -> Dict:
        """Process subjects dataframe."""
        subjects = {}
        for _, row in df.iterrows():
            subject_id = str(row.get('ID', row.get('Subject ID', '')))
            if subject_id:
                subjects[subject_id] = {
                    'name': row.get('Name', ''),
                    'hours_per_week': int(row.get('Hours/Week', 3)),
                    'requires_lab': bool(row.get('Requires Lab', False)),
                    'department': row.get('Department', '')
                }
        return subjects
        
    def _process_timeslots(self, df: pd.DataFrame) -> list:
        """Process timeslots dataframe."""
        timeslots = []
        for _, row in df.iterrows():
            timeslots.append({
                'day': row.get('Day', ''),
                'period': int(row.get('Period', 1)),
                'start_time': str(row.get('Start Time', '')),
                'end_time': str(row.get('End Time', ''))
            })
        return timeslots


# ============================================
# importer/template_importer.py
# ============================================
"""
Template-based importer for clean data entry
"""

import pandas as pd
import json
from pathlib import Path
from typing import Dict, Any


class TemplateImporter:
    """Import timetable data from clean template files."""
    
    def import_folder(self, folder_path: str) -> Dict[str, Any]:
        """Import data from template folder."""
        folder = Path(folder_path)
        data = {}
        
        # Import each template file
        if (folder / 'teachers.xlsx').exists():
            data['teachers'] = self._import_teachers(folder / 'teachers.xlsx')
            
        if (folder / 'classes.xlsx').exists():
            data['classes'] = self._import_classes(folder / 'classes.xlsx')
            
        if (folder / 'rooms.xlsx').exists():
            data['rooms'] = self._import_rooms(folder / 'rooms.xlsx')
            
        if (folder / 'subjects.xlsx').exists():
            data['subjects'] = self._import_subjects(folder / 'subjects.xlsx')
            
        if (folder / 'timeslots.xlsx').exists():
            data['timeslots'] = self._import_timeslots(folder / 'timeslots.xlsx')
            
        if (folder / 'constraints.json').exists():
            with open(folder / 'constraints.json', 'r') as f:
                data['constraints'] = json.load(f)
                
        return data
        
    def export_blank_templates(self, folder_path: str):
        """Export blank template files."""
        folder = Path(folder_path)
        folder.mkdir(parents=True, exist_ok=True)
        
        # Create teachers template
        teachers_df = pd.DataFrame({
            'Teacher ID': ['T001', 'T002'],
            'Name': ['John Smith', 'Jane Doe'],
            'Subjects': ['MATH,PHY', 'ENG,HIST'],
            'Max Hours': [20, 18],
            'Email': ['john@school.edu', 'jane@school.edu'],
            'Department': ['Science', 'Arts']
        })
        teachers_df.to_excel(folder / 'teachers.xlsx', index=False)
        
        # Create classes template
        classes_df = pd.DataFrame({
            'Class ID': ['C001', 'C002'],
            'Name': ['Grade 10A', 'Grade 10B'],
            'Size': [30, 28],
            'Level': ['Grade 10', 'Grade 10'],
            'Homeroom Teacher': ['T001', 'T002']
        })
        classes_df.to_excel(folder / 'classes.xlsx', index=False)
        
        # Create rooms template
        rooms_df = pd.DataFrame({
            'Room ID': ['R001', 'R002', 'LAB01'],
            'Name': ['Classroom 1', 'Classroom 2', 'Science Lab'],
            'Capacity': [35, 35, 30],
            'Type': ['Classroom', 'Classroom', 'Laboratory'],
            'Facilities': ['Projector,Whiteboard', 'Projector,Whiteboard', 'Lab Equipment']
        })
        rooms_df.to_excel(folder / 'rooms.xlsx', index=False)
        
        # Create subjects template
        subjects_df = pd.DataFrame({
            'Subject ID': ['MATH', 'ENG', 'PHY'],
            'Name': ['Mathematics', 'English', 'Physics'],
            'Hours/Week': [4, 4, 3],
            'Requires Lab': [False, False, True],
            'Department': ['Science', 'Arts', 'Science']
        })
        subjects_df.to_excel(folder / 'subjects.xlsx', index=False)
        
        # Create timeslots template
        timeslots_df = pd.DataFrame({
            'Day': ['Monday'] * 8 + ['Tuesday'] * 8 + ['Wednesday'] * 8 + 
                   ['Thursday'] * 8 + ['Friday'] * 8,
            'Period': list(range(1, 9)) * 5,
            'Start Time': ['08:00', '09:00', '10:00', '11:00', '12:00', '13:00', '14:00', '15:00'] * 5,
            'End Time': ['09:00', '10:00', '11:00', '12:00', '13:00', '14:00', '15:00', '16:00'] * 5
        })
        timeslots_df.to_excel(folder / 'timeslots.xlsx', index=False)
        
        # Create constraints template
        constraints = {
            "hard": [
                {"type": "no_teacher_conflict", "enabled": True},
                {"type": "no_room_conflict", "enabled": True},
                {"type": "no_class_conflict", "enabled": True}
            ],
            "soft": [
                {"type": "minimize_gaps", "weight": 5},
                {"type": "preferred_times", "weight": 3}
            ]
        }
        with open(folder / 'constraints.json', 'w') as f:
            json.dump(constraints, f, indent=2)
            
    def _import_teachers(self, file_path: Path) -> Dict:
        """Import teachers from template."""
        df = pd.read_excel(file_path)
        teachers = {}
        for _, row in df.iterrows():
            teacher_id = str(row['Teacher ID'])
            teachers[teacher_id] = {
                'name': row['Name'],
                'subjects': row['Subjects'].split(','),
                'max_hours': row['Max Hours'],
                'email': row.get('Email', ''),
                'department': row.get('Department', '')
            }
        return teachers
        
    def _import_classes(self, file_path: Path) -> Dict:
        """Import classes from template."""
        df = pd.read_excel(file_path)
        classes = {}
        for _, row in df.iterrows():
            class_id = str(row['Class ID'])
            classes[class_id] = {
                'name': row['Name'],
                'size': row['Size'],
                'level': row['Level'],
                'homeroom_teacher': row.get('Homeroom Teacher', '')
            }
        return classes
        
    def _import_rooms(self, file_path: Path) -> Dict:
        """Import rooms from template."""
        df = pd.read_excel(file_path)
        rooms = {}
        for _, row in df.iterrows():
            room_id = str(row['Room ID'])
            rooms[room_id] = {
                'name': row['Name'],
                'capacity': row['Capacity'],
                'type': row['Type'],
                'facilities': row.get('Facilities', '').split(',')
            }
        return rooms
        
    def _import_subjects(self, file_path: Path) -> Dict:
        """Import subjects from template."""
        df = pd.read_excel(file_path)
        subjects = {}
        for _, row in df.iterrows():
            subject_id = str(row['Subject ID'])
            subjects[subject_id] = {
                'name': row['Name'],
                'hours_per_week': row['Hours/Week'],
                'requires_lab': row.get('Requires Lab', False),
                'department': row.get('Department', '')
            }
        return subjects
        
    def _import_timeslots(self, file_path: Path) -> list:
        """Import timeslots from template."""
        df = pd.read_excel(file_path)
        timeslots = []
        for _, row in df.iterrows():
            timeslots.append({
                'day': row['Day'],
                'period': row['Period'],
                'start_time': str(row['Start Time']),
                'end_time': str(row['End Time'])
            })
        return timeslots


# ============================================
# exporter/excel_exporter.py
# ============================================
"""
Excel exporter for timetable solutions
"""

import pandas as pd
from pathlib import Path
from datetime import datetime
from typing import Dict, Any


class ExcelExporter:
    """Export timetable solution to Excel format."""
    
    def export(self, solution: Dict, folder: str, prefix: str, 
               include_stats: bool = True, include_conflicts: bool = True) -> str:
        """Export solution to Excel file."""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = Path(folder) / f"{prefix}_{timestamp}.xlsx"
        
        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            # Export main timetable
            self._export_timetable_by_class(solution, writer)
            self._export_timetable_by_teacher(solution, writer)
            self._export_timetable_by_room(solution, writer)
            
            # Export statistics
            if include_stats:
                self._export_statistics(solution, writer)
                
            # Export conflicts
            if include_conflicts:
                self._export_conflicts(solution, writer)
                
        return str(filename)
        
    def _export_timetable_by_class(self, solution: Dict, writer):
        """Export timetable organized by class."""
        assignments = solution.get('assignments', [])
        
        # Group by class
        class_schedules = {}
        for assignment in assignments:
            class_id = assignment['class']
            if class_id not in class_schedules:
                class_schedules[class_id] = []
            class_schedules[class_id].append(assignment)
            
        # Create dataframe for each class
        for class_id, schedule in class_schedules.items():
            df = self._create_timetable_df(schedule)
            df.to_excel(writer, sheet_name=f'Class_{class_id}', index=False)
            
    def _export_timetable_by_teacher(self, solution: Dict, writer):
        """Export timetable organized by teacher."""
        assignments = solution.get('assignments', [])
        
        # Group by teacher
        teacher_schedules = {}
        for assignment in assignments:
            teacher_id = assignment['teacher']
            if teacher_id not in teacher_schedules:
                teacher_schedules[teacher_id] = []
            teacher_schedules[teacher_id].append(assignment)
            
        # Create summary dataframe
        summary_data = []
        for teacher_id, schedule in teacher_schedules.items():
            df = self._create_timetable_df(schedule)
            summary_data.append({
                'Teacher': teacher_id,
                'Total Hours': len(schedule),
                'Classes': len(set(a['class'] for a in schedule)),
                'Rooms': len(set(a['room'] for a in schedule))
            })
            
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, sheet_name='Teacher_Summary', index=False)
        
    def _export_timetable_by_room(self, solution: Dict, writer):
        """Export timetable organized by room."""
        assignments = solution.get('assignments', [])
        
        # Group by room
        room_schedules = {}
        for assignment in assignments:
            room_id = assignment['room']
            if room_id not in room_schedules:
                room_schedules[room_id] = []
            room_schedules[room_id].append(assignment)
            
        # Create utilization dataframe
        util_data = []
        for room_id, schedule in room_schedules.items():
            util_data.append({
                'Room': room_id,
                'Total Hours': len(schedule),
                'Utilization %': (len(schedule) / 40) * 100  # Assuming 40 slots per week
            })
            
        util_df = pd.DataFrame(util_data)
        util_df.to_excel(writer, sheet_name='Room_Utilization', index=False)
        
    def _create_timetable_df(self, schedule: list) -> pd.DataFrame:
        """Create timetable dataframe from schedule."""
        # This would create a proper timetable grid
        # For now, returning a simple list
        data = []
        for assignment in schedule:
            data.append({
                'Day': self._get_day_from_slot(assignment['timeslot']),
                'Period': self._get_period_from_slot(assignment['timeslot']),
                'Subject': assignment['subject'],
                'Teacher': assignment['teacher'],
                'Room': assignment['room'],
                'Class': assignment['class']
            })
        return pd.DataFrame(data)
        
    def _get_day_from_slot(self, slot_id: int) -> str:
        """Get day name from slot ID."""
        days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']
        return days[slot_id // 8]
        
    def _get_period_from_slot(self, slot_id: int) -> int:
        """Get period number from slot ID."""
        return (slot_id % 8) + 1
        
    def _export_statistics(self, solution: Dict, writer):
        """Export solution statistics."""
        stats = solution.get('statistics', {})
        stats_df = pd.DataFrame([stats])
        stats_df.to_excel(writer, sheet_name='Statistics', index=False)
        
    def _export_conflicts(self, solution: Dict, writer):
        """Export conflict report."""
        conflicts = solution.get('conflicts', [])
        if conflicts:
            conflicts_df = pd.DataFrame(conflicts)
            conflicts_df.to_excel(writer, sheet_name='Conflicts', index=False)


# ============================================
# exporter/pdf_exporter.py
# ============================================
"""
PDF exporter for timetable solutions
"""

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
from pathlib import Path
from datetime import datetime
from typing import Dict


class PDFExporter:
    """Export timetable solution to PDF format."""
    
    def export(self, solution: Dict, folder: str, prefix: str, 
               include_stats: bool = True) -> str:
        """Export solution to PDF file."""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = Path(folder) / f"{prefix}_{timestamp}.pdf"
        
        doc = SimpleDocTemplate(
            str(filename),
            pagesize=landscape(A4),
            rightMargin=30,
            leftMargin=30,
            topMargin=30,
            bottomMargin=30
        )
        
        elements = []
        styles = getSampleStyleSheet()
        
        # Title
        title = Paragraph("Timetable Schedule", styles['Title'])
        elements.append(title)
        elements.append(Spacer(1, 0.5*inch))
        
        # Add timetables
        self._add_class_timetables(solution, elements, styles)
        
        # Add statistics
        if include_stats:
            self._add_statistics(solution, elements, styles)
            
        # Build PDF
        doc.build(elements)
        
        return str(filename)
        
    def _add_class_timetables(self, solution: Dict, elements, styles):
        """Add class timetables to PDF."""
        assignments = solution.get('assignments', [])
        
        # Group by class
        class_schedules = {}
        for assignment in assignments:
            class_id = assignment['class']
            if class_id not in class_schedules:
                class_schedules[class_id] = []
            class_schedules[class_id].append(assignment)
            
        # Create table for each class
        for class_id, schedule in class_schedules.items():
            # Class header
            header = Paragraph(f"Class: {class_id}", styles['Heading2'])
            elements.append(header)
            elements.append(Spacer(1, 0.2*inch))
            
            # Create timetable grid
            table_data = self._create_timetable_grid(schedule)
            table = Table(table_data)
            
            # Style the table
            table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 10),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                ('GRID', (0, 0), (-1, -1), 1, colors.black)
            ]))
            
            elements.append(table)
            elements.append(Spacer(1, 0.5*inch))
            
    def _create_timetable_grid(self, schedule: list) -> list:
        """Create timetable grid data."""
        # Initialize grid
        days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']
        periods = 8
        
        grid = [['Time'] + days]
        
        for period in range(1, periods + 1):
            row = [f'Period {period}']
            for day in days:
                # Find assignment for this slot
                cell_content = ''
                for assignment in schedule:
                    if (self._get_day_from_slot(assignment['timeslot']) == day and
                        self._get_period_from_slot(assignment['timeslot']) == period):
                        cell_content = f"{assignment['subject']}\n{assignment['teacher']}\n{assignment['room']}"
                        break
                row.append(cell_content)
            grid.append(row)
            
        return grid
        
    def _get_day_from_slot(self, slot_id: int) -> str:
        """Get day name from slot ID."""
        days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']
        return days[slot_id // 8]
        
    def _get_period_from_slot(self, slot_id: int) -> int:
        """Get period number from slot ID."""
        return (slot_id % 8) + 1
        
    def _add_statistics(self, solution: Dict, elements, styles):
        """Add statistics to PDF."""
        stats = solution.get('statistics', {})
        
        # Statistics header
        header = Paragraph("Statistics", styles['Heading2'])
        elements.append(header)
        elements.append(Spacer(1, 0.2*inch))
        
        # Create statistics table
        stats_data = [
            ['Metric', 'Value'],
            ['Total Assignments', str(stats.get('total_assignments', 0))],
            ['Conflicts', str(stats.get('conflicts', 0))],
            ['Room Utilization', f"{stats.get('utilization', 0):.1f}%"],
            ['Solve Time', f"{stats.get('solve_time', 0):.2f}s"]
        ]
        
        stats_table = Table(stats_data)
        stats_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]))
        
        elements.append(stats_table)


# ============================================
# constraints/constraint_editor.py
# ============================================
"""
Constraint editor widget for configuring solver constraints
"""

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QGroupBox,
    QListWidget, QListWidgetItem, QPushButton,
    QCheckBox, QSpinBox, QLabel, QComboBox,
    QDialog, QDialogButtonBox, QFormLayout,
    QTextEdit, QMessageBox
)
from PySide6.QtCore import Qt
import json
from typing import Dict, Any


class ConstraintEditor(QWidget):
    """Widget for editing timetable constraints."""
    
    def __init__(self):
        super().__init__()
        self.constraints = {
            'hard': [],
            'soft': []
        }
        self.data = None
        self.init_ui()
        self.load_default_constraints()
        
    def init_ui(self):
        """Initialize the constraint editor UI."""
        layout = QVBoxLayout()
        
        # Hard constraints
        hard_group = QGroupBox("Hard Constraints (Must be satisfied)")
        hard_layout = QVBoxLayout()
        
        self.hard_list = QListWidget()
        hard_layout.addWidget(self.hard_list)
        
        hard_btn_layout = QHBoxLayout()
        add_hard_btn = QPushButton("Add Hard Constraint")
        add_hard_btn.clicked.connect(lambda: self.add_constraint('hard'))
        hard_btn_layout.addWidget(add_hard_btn)
        
        edit_hard_btn = QPushButton("Edit")
        edit_hard_btn.clicked.connect(lambda: self.edit_constraint('hard'))
        hard_btn_layout.addWidget(edit_hard_btn)
        
        remove_hard_btn = QPushButton("Remove")
        remove_hard_btn.clicked.connect(lambda: self.remove_constraint('hard'))
        hard_btn_layout.addWidget(remove_hard_btn)
        
        hard_layout.addLayout(hard_btn_layout)
        hard_group.setLayout(hard_layout)
        layout.addWidget(hard_group)
        
        # Soft constraints
        soft_group = QGroupBox("Soft Constraints (Preferences)")
        soft_layout = QVBoxLayout()
        
        self.soft_list = QListWidget()
        soft_layout.addWidget(self.soft_list)
        
        soft_btn_layout = QHBoxLayout()
        add_soft_btn = QPushButton("Add Soft Constraint")
        add_soft_btn.clicked.connect(lambda: self.add_constraint('soft'))
        soft_btn_layout.addWidget(add_soft_btn)
        
        edit_soft_btn = QPushButton("Edit")
        edit_soft_btn.clicked.connect(lambda: self.edit_constraint('soft'))
        soft_btn_layout.addWidget(edit_soft_btn)
        
        remove_soft_btn = QPushButton("Remove")
        remove_soft_btn.clicked.connect(lambda: self.remove_constraint('soft'))
        soft_btn_layout.addWidget(remove_soft_btn)
        
        soft_layout.addLayout(soft_btn_layout)
        soft_group.setLayout(soft_layout)
        layout.addWidget(soft_group)
        
        # Constraint templates
        template_group = QGroupBox("Quick Templates")
        template_layout = QHBoxLayout()
        
        load_basic_btn = QPushButton("Load Basic Set")
        load_basic_btn.clicked.connect(self.load_basic_template)
        template_layout.addWidget(load_basic_btn)
        
        load_strict_btn = QPushButton("Load Strict Set")
        load_strict_btn.clicked.connect(self.load_strict_template)
        template_layout.addWidget(load_strict_btn)
        
        export_btn = QPushButton("Export Constraints")
        export_btn.clicked.connect(self.export_constraints)
        template_layout.addWidget(export_btn)
        
        import_btn = QPushButton("Import Constraints")
        import_btn.clicked.connect(self.import_constraints)
        template_layout.addWidget(import_btn)
        
        template_group.setLayout(template_layout)
        layout.addWidget(template_group)
        
        self.setLayout(layout)
        
    def load_default_constraints(self):
        """Load default constraints."""
        self.constraints['hard'] = [
            {'type': 'no_teacher_conflict', 'name': 'No Teacher Conflicts', 'enabled': True},
            {'type': 'no_room_conflict', 'name': 'No Room Conflicts', 'enabled': True},
            {'type': 'no_class_conflict', 'name': 'No Class Conflicts', 'enabled': True},
            {'type': 'room_capacity', 'name': 'Respect Room Capacity', 'enabled': True}
        ]
        
        self.constraints['soft'] = [
            {'type': 'minimize_gaps', 'name': 'Minimize Teacher Gaps', 'weight': 5},
            {'type': 'preferred_times', 'name': 'Respect Preferred Times', 'weight': 3}
        ]
        
        self.update_constraint_lists()
        
    def update_constraint_lists(self):
        """Update the constraint list widgets."""
        self.hard_list.clear()
        for constraint in self.constraints['hard']:
            item_text = f"[{'✓' if constraint.get('enabled', True) else '✗'}] {constraint['name']}"
            self.hard_list.addItem(item_text)
            
        self.soft_list.clear()
        for constraint in self.constraints['soft']:
            item_text = f"{constraint['name']} (Weight: {constraint.get('weight', 1)})"
            self.soft_list.addItem(item_text)
            
    def add_constraint(self, constraint_type):
        """Add a new constraint."""
        dialog = ConstraintDialog(constraint_type, None, self.data, self)
        if dialog.exec():
            constraint = dialog.get_constraint()
            self.constraints[constraint_type].append(constraint)
            self.update_constraint_lists()
            
    def edit_constraint(self, constraint_type):
        """Edit selected constraint."""
        if constraint_type == 'hard':
            list_widget = self.hard_list
        else:
            list_widget = self.soft_list
            
        current_row = list_widget.currentRow()
        if current_row >= 0:
            constraint = self.constraints[constraint_type][current_row]
            dialog = ConstraintDialog(constraint_type, constraint, self.data, self)
            if dialog.exec():
                self.constraints[constraint_type][current_row] = dialog.get_constraint()
                self.update_constraint_lists()
                
    def remove_constraint(self, constraint_type):
        """Remove selected constraint."""
        if constraint_type == 'hard':
            list_widget = self.hard_list
        else:
            list_widget = self.soft_list
            
        current_row = list_widget.currentRow()
        if current_row >= 0:
            del self.constraints[constraint_type][current_row]
            self.update_constraint_lists()
            
    def load_basic_template(self):
        """Load basic constraint template."""
        self.constraints['hard'] = [
            {'type': 'no_teacher_conflict', 'name': 'No Teacher Conflicts', 'enabled': True},
            {'type': 'no_room_conflict', 'name': 'No Room Conflicts', 'enabled': True},
            {'type': 'no_class_conflict', 'name': 'No Class Conflicts', 'enabled': True}
        ]
        
        self.constraints['soft'] = [
            {'type': 'minimize_gaps', 'name': 'Minimize Teacher Gaps', 'weight': 3}
        ]
        
        self.update_constraint_lists()
        QMessageBox.information(self, "Template Loaded", "Basic constraint set loaded.")
        
    def load_strict_template(self):
        """Load strict constraint template."""
        self.constraints['hard'] = [
            {'type': 'no_teacher_conflict', 'name': 'No Teacher Conflicts', 'enabled': True},
            {'type': 'no_room_conflict', 'name': 'No Room Conflicts', 'enabled': True},
            {'type': 'no_class_conflict', 'name': 'No Class Conflicts', 'enabled': True},
            {'type': 'room_capacity', 'name': 'Respect Room Capacity', 'enabled': True},
            {'type': 'teacher_availability', 'name': 'Respect Teacher Availability', 'enabled': True},
            {'type': 'max_hours_per_day', 'name': 'Max Hours Per Day', 'enabled': True, 'max_hours': 6}
        ]
        
        self.constraints['soft'] = [
            {'type': 'minimize_gaps', 'name': 'Minimize Teacher Gaps', 'weight': 10},
            {'type': 'preferred_times', 'name': 'Respect Preferred Times', 'weight': 8},
            {'type': 'balance_workload', 'name': 'Balance Teacher Workload', 'weight': 5}
        ]
        
        self.update_constraint_lists()
        QMessageBox.information(self, "Template Loaded", "Strict constraint set loaded.")
        
    def export_constraints(self):
        """Export constraints to file."""
        from PySide6.QtWidgets import QFileDialog
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Export Constraints",
            "constraints.json",
            "JSON Files (*.json)"
        )
        
        if file_path:
            with open(file_path, 'w') as f:
                json.dump(self.constraints, f, indent=2)
            QMessageBox.information(self, "Export Success", f"Constraints exported to {file_path}")
            
    def import_constraints(self):
        """Import constraints from file."""
        from PySide6.QtWidgets import QFileDialog
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Import Constraints",
            "",
            "JSON Files (*.json)"
        )
        
        if file_path:
            try:
                with open(file_path, 'r') as f:
                    self.constraints = json.load(f)
                self.update_constraint_lists()
                QMessageBox.information(self, "Import Success", "Constraints imported successfully.")
            except Exception as e:
                QMessageBox.critical(self, "Import Error", f"Failed to import constraints: {str(e)}")
                
    def get_constraints(self) -> Dict:
        """Get current constraints."""
        return self.constraints
        
    def load_constraints(self, constraints: Dict):
        """Load constraints from dict."""
        self.constraints = constraints
        self.update_constraint_lists()
        
    def load_data(self, data: Dict):
        """Load timetable data for constraint configuration."""
        self.data = data
        
    def clear(self):
        """Clear all constraints."""
        self.constraints = {'hard': [], 'soft': []}
        self.update_constraint_lists()


class ConstraintDialog(QDialog):
    """Dialog for adding/editing constraints."""
    
    def __init__(self, constraint_type: str, constraint: Dict = None, 
                 data: Dict = None, parent=None):
        super().__init__(parent)
        self.constraint_type = constraint_type
        self.constraint = constraint or {}
        self.data = data
        self.init_ui()
        
    def init_ui(self):
        """Initialize dialog UI."""
        self.setWindowTitle(f"{'Edit' if self.constraint else 'Add'} {self.constraint_type.title()} Constraint")
        self.setModal(True)
        
        layout = QFormLayout()
        
        # Constraint type selector
        self.type_combo = QComboBox()
        if self.constraint_type == 'hard':
            self.type_combo.addItems([
                'no_teacher_conflict',
                'no_room_conflict',
                'no_class_conflict',
                'room_capacity',
                'teacher_availability',
                'max_hours_per_day',
                'subject_continuity'
            ])
        else:
            self.type_combo.addItems([
                'minimize_gaps',
                'preferred_times',
                'balance_workload',
                'minimize_room_changes',
                'group_subjects'
            ])
            
        if self.constraint:
            self.type_combo.setCurrentText(self.constraint.get('type', ''))
            
        layout.addRow("Type:", self.type_combo)
        
        # Constraint name
        self.name_edit = QLineEdit(self.constraint.get('name', ''))
        layout.addRow("Name:", self.name_edit)
        
        # Enabled checkbox (for hard constraints)
        if self.constraint_type == 'hard':
            self.enabled_check = QCheckBox()
            self.enabled_check.setChecked(self.constraint.get('enabled', True))
            layout.addRow("Enabled:", self.enabled_check)
            
        # Weight spinbox (for soft constraints)
        if self.constraint_type == 'soft':
            self.weight_spin = QSpinBox()
            self.weight_spin.setRange(1, 10)
            self.weight_spin.setValue(self.constraint.get('weight', 5))
            layout.addRow("Weight:", self.weight_spin)
            
        # Additional parameters based on type
        self.params_widget = QWidget()
        self.params_layout = QFormLayout(self.params_widget)
        layout.addRow("Parameters:", self.params_widget)
        
        # Update parameters when type changes
        self.type_combo.currentTextChanged.connect(self.update_parameters)
        self.update_parameters()
        
        # Buttons
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addRow(buttons)
        
        self.setLayout(layout)
        
    def update_parameters(self):
        """Update parameter fields based on selected type."""
        # Clear existing parameters
        while self.params_layout.rowCount() > 0:
            self.params_layout.removeRow(0)
            
        constraint_type = self.type_combo.currentText()
        
        if constraint_type == 'max_hours_per_day':
            self.max_hours_spin = QSpinBox()
            self.max_hours_spin.setRange(1, 8)
            self.max_hours_spin.setValue(self.constraint.get('max_hours', 6))
            self.params_layout.addRow("Max Hours:", self.max_hours_spin)
            
        elif constraint_type == 'teacher_availability':
            self.availability_text = QTextEdit()
            self.availability_text.setPlainText(
                json.dumps(self.constraint.get('data', {}), indent=2)
            )
            self.params_layout.addRow("Availability Data:", self.availability_text)
            
    def get_constraint(self) -> Dict:
        """Get configured constraint."""
        constraint = {
            'type': self.type_combo.currentText(),
            'name': self.name_edit.text() or self.type_combo.currentText()
        }
        
        if self.constraint_type == 'hard':
            constraint['enabled'] = self.enabled_check.isChecked()
        else:
            constraint['weight'] = self.weight_spin.value()
            
        # Add type-specific parameters
        if constraint['type'] == 'max_hours_per_day':
            constraint['max_hours'] = self.max_hours_spin.value()
        elif constraint['type'] == 'teacher_availability':
            try:
                constraint['data'] = json.loads(self.availability_text.toPlainText())
            except:
                constraint['data'] = {}
                
        return constraint


# ============================================
# utils/validation.py
# ============================================
"""
Validation utilities for timetable data and constraints
"""

from typing import Dict, List, Any


class ValidationReport:
    """Generate validation reports for timetable data."""
    
    def __init__(self):
        self.errors = []
        self.warnings = []
        self.info = []
        
    def validate(self, data: Dict, constraints: Dict) -> 'ValidationReport':
        """Validate data and constraints."""
        self.errors = []
        self.warnings = []
        self.info = []
        
        # Validate data completeness
        self._validate_data_completeness(data)
        
        # Validate data consistency
        self._validate_data_consistency(data)
        
        # Validate constraints
        self._validate_constraints(constraints, data)
        
        return self
        
    def _validate_data_completeness(self, data: Dict):
        """Check if all required data is present."""
        required_keys = ['teachers', 'classes', 'rooms', 'subjects', 'timeslots']
        
        for key in required_keys:
            if key not in data or not data[key]:
                self.errors.append(f"Missing required data: {key}")
            else:
                self.info.append(f"{key.title()}: {len(data[key])} items")
                
    def _validate_data_consistency(self, data: Dict):
        """Check data consistency."""
        # Check teacher subjects exist
        if 'teachers' in data and 'subjects' in data:
            all_subjects = set(data['subjects'].keys())
            for teacher_id, teacher in data['teachers'].items():
                for subject in teacher.get('subjects', []):
                    if subject and subject not in all_subjects:
                        self.warnings.append(
                            f"Teacher {teacher_id} teaches unknown subject: {subject}"
                        )
                        
        # Check room capacity vs class size
        if 'rooms' in data and 'classes' in data:
            for class_id, class_data in data['classes'].items():
                class_size = class_data.get('size', 0)
                suitable_rooms = 0
                for room_id, room in data['rooms'].items():
                    if room.get('capacity', 0) >= class_size:
                        suitable_rooms += 1
                        
                if suitable_rooms == 0:
                    self.errors.append(
                        f"No room can accommodate class {class_id} (size: {class_size})"
                    )
                elif suitable_rooms < 3:
                    self.warnings.append(
                        f"Only {suitable_rooms} rooms can accommodate class {class_id}"
                    )
                    
    def _validate_constraints(self, constraints: Dict, data: Dict):
        """Validate constraints against data."""
        hard_constraints = constraints.get('hard', [])
        soft_constraints = constraints.get('soft', [])
        
        # Check if essential hard constraints are present
        essential = ['no_teacher_conflict', 'no_room_conflict', 'no_class_conflict']
        present = [c['type'] for c in hard_constraints if c.get('enabled', True)]
        
        for constraint_type in essential:
            if constraint_type not in present:
                self.warnings.append(f"Essential constraint not enabled: {constraint_type}")
                
        # Check constraint feasibility
        total_slots = len(data.get('timeslots', []))
        total_classes = len(data.get('classes', {}))
        total_rooms = len(data.get('rooms', {}))
        
        max_possible_assignments = min(total_slots * total_rooms, total_slots * total_classes)
        
        self.info.append(f"Maximum possible assignments: {max_possible_assignments}")
        
    def to_string(self) -> str:
        """Convert report to string."""
        report = "=== VALIDATION REPORT ===\n\n"
        
        if self.errors:
            report += "ERRORS:\n"
            for error in self.errors:
                report += f"  ✗ {error}\n"
            report += "\n"
            
        if self.warnings:
            report += "WARNINGS:\n"
            for warning in self.warnings:
                report += f"  ⚠ {warning}\n"
            report += "\n"
            
        if self.info:
            report += "INFORMATION:\n"
            for info in self.info:
                report += f"  ℹ {info}\n"
            report += "\n"
            
        if not self.errors and not self.warnings:
            report += "✓ All validations passed successfully!\n"
            
        return report
        
    def has_errors(self) -> bool:
        """Check if validation has errors."""
        return len(self.errors) > 0
        
    def has_warnings(self) -> bool:
        """Check if validation has warnings."""
        return len(self.warnings) > 0


# ============================================
# utils/settings.py
# ============================================
"""
Settings management for the application
"""

from PySide6.QtCore import QSettings
from pathlib import Path
import os


class SettingsManager:
    """Manage application settings."""
    
    def __init__(self):
        self.settings = QSettings("Peninsula College", "Timetable Planner")
        self._init_defaults()
        
    def _init_defaults(self):
        """Initialize default settings."""
        if not self.settings.contains("export_folder"):
            default_path = str(Path.home() / "OneDrive - PENINSULA COLLEGE (NORTHERN) SDN BHD" / 
                             "Documents" / "01. PEEC" / "01. PYTHON" / "Timetable" / "202509")
            
            # Use short path if default is too long or doesn't exist
            if len(default_path) > 200 or not Path(default_path).parent.exists():
                default_path = "C:\\Apps\\Timetable\\202509"
                
            self.settings.setValue("export_folder", default_path)
            
    def get_export_folder(self) -> str:
        """Get export folder path."""
        path = self.settings.value("export_folder", "")
        
        # Create folder if it doesn't exist
        if path:
            Path(path).mkdir(parents=True, exist_ok=True)
            
        return path
        
    def set_export_folder(self, path: str):
        """Set export folder path."""
        self.settings.setValue("export_folder", path)
        
    def get_last_import_folder(self) -> str:
        """Get last import folder."""
        return self.settings.value("last_import_folder", str(Path.home()))
        
    def set_last_import_folder(self, path: str):
        """Set last import folder."""
        self.settings.setValue("last_import_folder", path)
        
    def get_last_project_folder(self) -> str:
        """Get last project folder."""
        return self.settings.value("last_project_folder", str(Path.home()))
        
    def set_last_project_folder(self, path: str):
        """Set last project folder."""
        self.settings.setValue("last_project_folder", path)
        
    def get_use_short_path(self) -> bool:
        """Get whether to use short path fallback."""
        return self.settings.value("use_short_path", False, type=bool)
        
    def set_use_short_path(self, use_short: bool):
        """Set whether to use short path fallback."""
        self.settings.setValue("use_short_path", use_short)
        
        if use_short:
            self.set_export_folder("C:\\Apps\\Timetable\\202509")# ============================================
# solver/timetable_solver.py
# ============================================
"""
Timetable solver using OR-Tools CP-SAT solver
"""

from ortools.sat.python import cp_model
from typing import Dict, List, Optional, Callable, Any
import time


class TimetableSolver:
    """Constraint Programming solver for timetable generation."""
    
    def __init__(self, params: Dict[str, Any]):
        self.data = params['data']
        self.constraints = params['constraints']
        self.time_limit = params.get('time_limit', 60)
        self.num_workers = params.get('num_workers', 4)
        self.search_strategy = params.get('search_strategy', 'Automatic')
        self.progress_callback = None
        self.model = cp_model.CpModel()
        self.solver = cp_model.CpSolver()
        
    def set_progress_callback(self, callback: Callable[[int, str], None]):
        """Set callback for progress updates."""
        self.progress_callback = callback
        
    def _report_progress(self, percent: int, message: str):
        """Report progress if callback is set."""
        if self.progress_callback:
            self.progress_callback(percent, message)
            
    def solve(self) -> Optional[Dict]:
        """Solve the timetable problem."""
        try:
            self._report_progress(0, "Building model...")
            self._build_model()
            
            self._report_progress(20, "Adding constraints...")
            self._add_constraints()
            
            self._report_progress(40, "Configuring solver...")
            self._configure_solver()
            
            self._report_progress(50, "Running solver...")
            status = self.solver.Solve(self.model)
            
            if status == cp_model.OPTIMAL or status == cp_model.FEASIBLE:
                self._report_progress(90, "Extracting solution...")
                solution = self._extract_solution()
                self._report_progress(100, "Solution found!")
                return solution
            else:
                self._report_progress(100, "No feasible solution found")
                return None
                
        except Exception as e:
            raise Exception(f"Solver error: {str(e)}")
            
    def _build_model(self):
        """Build the CP model with decision variables."""
        # Extract data
        teachers = self.data.get('teachers', {})
        classes = self.data.get('classes', {})
        rooms = self.data.get('rooms', {})
        subjects = self.data.get('subjects', {})
        timeslots = self.data.get('timeslots', [])
        
        # Create decision variables
        self.assignments = {}
        
        for teacher_id in teachers:
            for class_id in classes:
                for room_id in rooms:
                    for slot_id, slot in enumerate(timeslots):
                        for subject_id in subjects:
                            var_name = f"t{teacher_id}_c{class_id}_r{room_id}_s{slot_id}_sub{subject_id}"
                            self.assignments[
                                (teacher_id, class_id, room_id, slot_id, subject_id)
                            ] = self.model.NewBoolVar(var_name)
                            
    def _add_constraints(self):
        """Add constraints to the model."""
        # Hard constraints from configuration
        hard_constraints = self.constraints.get('hard', [])
        
        for constraint in hard_constraints:
            if constraint['type'] == 'no_teacher_conflict':
                self._add_no_teacher_conflict()
            elif constraint['type'] == 'no_room_conflict':
                self._add_no_room_conflict()
            elif constraint['type'] == 'no_class_conflict':
                self._add_no_class_conflict()
            elif constraint['type'] == 'teacher_availability':
                self._add_teacher_availability(constraint.get('data', {}))
            elif constraint['type'] == 'room_capacity':
                self._add_room_capacity()
                
        # Soft constraints (preferences)
        soft_constraints = self.constraints.get('soft', [])
        self.soft_penalties = []
        
        for constraint in soft_constraints:
            if constraint['type'] == 'minimize_gaps':
                self._add_minimize_gaps(constraint.get('weight', 1))
            elif constraint['type'] == 'preferred_times':
                self._add_preferred_times(constraint.get('data', {}), 
                                         constraint.get('weight', 1))
                                         
    def _add_no_teacher_conflict(self):
        """A teacher can only be in one place at a time."""
        teachers = self.data.get('teachers', {})
        timeslots = self.data.get('timeslots', [])
        
        for teacher_id in teachers:
            for slot_id, _ in enumerate(timeslots):
                vars_at_time = []
                for key, var in self.assignments.items():
                    if key[0] == teacher_id and key[3] == slot_id:
                        vars_at_time.append(var)
                        
                if vars_at_time:
                    self.model.Add(sum(vars_at_time) <= 1)time) <= 1)
                    
    def _add_no_room_conflict(self):
        """A room can only host one class at a time."""
        rooms = self.data.get('rooms', {})
        timeslots = self.data.get('timeslots', [])
        
        for room_id in rooms:
            for slot_id, _ in enumerate(timeslots):
                vars_at_time = []
                for key, var in self.assignments.items():
                    if key[2] == room_id and key[3] == slot_id:
                        vars_at_time.append(var)
                        
                if vars_at_time:
                    self.model.Add(sum(vars_at_time) <= 1)
                    
    def _add_no_class_conflict(self):
        """A class can only have one lesson at a time."""
        classes = self.data.get('classes', {})
        timeslots = self.data.get('timeslots', [])
        
        for class_id in classes:
            for slot_id, _ in enumerate(timeslots):
                vars_at_time = []
                for key, var in self.assignments.items():
                    if key[1] == class_id and key[3] == slot_id:
                        vars_at_time.append(var)
                        
                if vars_at_time:
                    self.model.Add(sum(vars_at_
