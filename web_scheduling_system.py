import pandas as pd
import pyodbc
from datetime import datetime, timedelta
import random
from collections import defaultdict
import sqlalchemy
from sqlalchemy import create_engine
import json
import os

# ===============================================
#  Database Configuration Macros
# ===============================================

# Modify these settings to match your environment


DB_SERVER = "localhost"
DB_NAME = "TestSchedulingDB"
DB_DRIVER = "ODBC Driver 17 for SQL Server"


# Database User Configuration - Currently using dedicated account
DB_USER = "SchedulingAppUser"
DB_PASSWORD = "SchedulingApp2025!"


# sa account configuration (backup)
# DB_USER = "sa"
# DB_PASSWORD = "Aa764144231!"

# For Windows Authentication, set to True and ensure service account has DB permissions
USE_WINDOWS_AUTH = False

# ===============================================

class WebSchedulingSystem:
    def __init__(self):
        # Database connections using configuration macros
        if USE_WINDOWS_AUTH:
            # Windows Authentication
            self.conn_string = (
                f"mssql+pyodbc://@{DB_SERVER}/{DB_NAME}?"
                f"driver={DB_DRIVER.replace(' ', '+')}&trusted_connection=yes"
            )
            self.pyodbc_conn_string = (
                f"Driver={{{DB_DRIVER}}};"
                f"Server={DB_SERVER};"
                f"Database={DB_NAME};"
                "Trusted_Connection=yes;"
            )
        else:
            # SQL Server Authentication
            self.conn_string = (
                f"mssql+pyodbc://{DB_USER}:{DB_PASSWORD}@{DB_SERVER}/{DB_NAME}?"
                f"driver={DB_DRIVER.replace(' ', '+')}"
            )
            self.pyodbc_conn_string = (
                f"Driver={{{DB_DRIVER}}};"
                f"Server={DB_SERVER};"
                f"Database={DB_NAME};"
                f"UID={DB_USER};"
                f"PWD={DB_PASSWORD};"
            )
        
        self.engine = create_engine(self.conn_string)
        self.conn = pyodbc.connect(self.pyodbc_conn_string)
        
        # Resource selection state
        self.selections = {
            'term': None,
            'session': None,
            'campus': None,
            'acad_group': None,
            'subject': None,
            'classes': [],
            'teachers': [],
            'rooms': []
        }
        
        # Available options cache
        self.available_options = {}
        
        # Time slot management - 
        self.disabled_time_slots = self._initialize_disabled_time_slots()
        
       
        self.current_schedule_results = {
            'scheduled_sessions': [],
            'conflicts': [],
            'generated': False,
            'timestamp': None
        }
        
        # Standard Excel columns (36 columns)
        self.standard_columns = [
            'Access', 'Term', 'Assign_Type', 'Class_Nbr', 'Offer_Nbr', 'Max_Units', 
            'Enrl_Stat', 'Long_Title', 'Component', 'Catalog', 'Acad_Group', 'Pat', 
            'Pat_Nbr', 'Session', 'F_ID', 'First_Name', 'Last_Name', 'Role', 'Career', 
            'Start_Date', 'End_Date', 'Course_ID', 'Course_Code', 'Subject', 'Descr', 
            'Section', 'Class_Stat', 'Mtg_Start', 'Mtg_End', 'Campus', 'Tot_Enrl', 
            'Cap_Enrl', 'Facil_ID', 'Day', 'Room_ID', 'Room_Capacity'
        ]
    
    def _initialize_disabled_time_slots(self):
        """Initialize disabled time slots, only keeping 7/49 time slots available"""
        days = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
        periods = [
            ("08:00", "09:15", "Period 1"),
            ("09:30", "10:45", "Period 2"),
            ("11:00", "12:15", "Period 3"),
            ("13:00", "14:15", "Period 4"),
            ("14:30", "15:45", "Period 5"),
            ("16:00", "17:15", "Period 6"),
            ("17:30", "18:45", "Period 7"),
        ]
        
        # Generate all time slot IDs
        all_time_slots = set()
        for day in days:
            for start, end, name in periods:
                time_id = f"{day}_{start}-{end}"
                all_time_slots.add(time_id)
        
        # Only keep 7 available time slots (e.g., Monday to Friday's first two periods, plus Wednesday's third period)
        available_slots = {
            "Monday_08:00-09:15",
            "Monday_09:30-10:45", 
            "Tuesday_08:00-09:15",
            "Wednesday_08:00-09:15",
            "Wednesday_11:00-12:15",
            "Thursday_08:00-09:15",
            "Friday_08:00-09:15"
        }
        
        # Return disabled time slots (49 total minus 7 available)
        disabled_slots = all_time_slots - available_slots
        return disabled_slots
    
    def get_temp_connection(self):
        """Create a temporary database connection using the same configuration"""
        return pyodbc.connect(self.pyodbc_conn_string)
    
    def load_available_resources(self):
        """Load all available resources for selection"""
        print("Loading available resources...")
        
        # 1. Load Terms
        terms_sql = "SELECT Term_Code, Term_Name, Session, Start_Date, End_Date FROM Term ORDER BY Term_Code"
        self.available_options['terms'] = pd.read_sql(terms_sql, self.engine)
        
        # 2. Load Campuses
        campus_sql = "SELECT Campus, Description FROM Campus ORDER BY Campus"
        self.available_options['campuses'] = pd.read_sql(campus_sql, self.engine)
        
        # 3. Load Academic Groups
        acad_group_sql = "SELECT DISTINCT Acad_Group FROM CourseCatalog WHERE Acad_Group IS NOT NULL ORDER BY Acad_Group"
        self.available_options['acad_groups'] = pd.read_sql(acad_group_sql, self.engine)
        
        # 4. Load Subjects (by Academic Group)
        subject_sql = """
        SELECT cc.Subject, cc.Acad_Group, COUNT(*) as Course_Count
        FROM CourseCatalog cc 
        WHERE cc.Subject IS NOT NULL 
        GROUP BY cc.Subject, cc.Acad_Group 
        ORDER BY cc.Acad_Group, cc.Subject
        """
        self.available_options['subjects'] = pd.read_sql(subject_sql, self.engine)
        
        # 5. Load Teachers
        teacher_sql = """
        SELECT t.F_ID, t.First_Name, t.Last_Name, 
               COUNT(DISTINCT ci.Class_Nbr) as Teaching_Load
        FROM Teacher t
        LEFT JOIN ClassInstructor ci ON t.F_ID = ci.F_ID
        GROUP BY t.F_ID, t.First_Name, t.Last_Name
        ORDER BY t.Last_Name, t.First_Name
        """
        self.available_options['teachers'] = pd.read_sql(teacher_sql, self.engine)
        
        # 6. Load Rooms
        room_sql = """
        SELECT Room_ID, Description, Capacity, Gender, Location, Facil_ID
        FROM Room 
        WHERE Capacity > 0
        ORDER BY Location, Capacity DESC
        """
        self.available_options['rooms'] = pd.read_sql(room_sql, self.engine)
        
        print("Resources loaded successfully!")
        return self.available_options
    
    def get_terms(self):
        """Get available terms"""
        if 'terms' not in self.available_options:
            self.load_available_resources()
        return self.available_options['terms'].to_dict('records')
    
    def get_sessions(self):
        """Get available sessions"""
        if 'terms' not in self.available_options:
            self.load_available_resources()
        sessions = self.available_options['terms']['Session'].unique()
        return [{'Session': s} for s in sessions]
    
    def get_campuses(self):
        """Get available campuses"""
        if 'campuses' not in self.available_options:
            self.load_available_resources()
        return self.available_options['campuses'].to_dict('records')
    
    def get_acad_groups(self):
        """Get available academic groups"""
        if 'acad_groups' not in self.available_options:
            self.load_available_resources()
        return self.available_options['acad_groups'].to_dict('records')
    
    def get_subjects_by_acad_group(self, acad_group):
        """Get subjects filtered by academic group with correct course count (distinct courses)"""
        if 'subjects' not in self.available_options:
            self.load_available_resources()
        
        # Use deduplicated course count statistics
        subject_sql = """
        SELECT cc.Subject, cc.Acad_Group, 
               COUNT(DISTINCT cc.Course_Code) as Course_Count
        FROM CourseCatalog cc 
        WHERE cc.Acad_Group = ? AND cc.Subject IS NOT NULL 
          AND cc.Course_Code IS NOT NULL AND cc.Course_Code != ''
        GROUP BY cc.Subject, cc.Acad_Group 
        ORDER BY cc.Subject
        """
        
        # Create new database connection to avoid concurrency conflicts
        temp_conn = self.get_temp_connection()
        
        try:
            cursor = temp_conn.cursor()
            cursor.execute(subject_sql, acad_group)
            
            columns = [col[0] for col in cursor.description]
            rows = cursor.fetchall()
            
            subjects_data = []
            for row in rows:
                subjects_data.append(dict(zip(columns, row)))
            
            return subjects_data
        finally:
            temp_conn.close()
    
    def get_classes_by_subject(self, subject):
        """Get unique courses (not classes) filtered by subject - returns Course_Code and Course_Title only"""
        courses_sql = """
        SELECT DISTINCT
            cc.Course_Code, cc.Course_Title
        FROM CourseCatalog cc
        WHERE cc.Subject = ? AND cc.Course_Code IS NOT NULL AND cc.Course_Code != ''
        ORDER BY cc.Course_Code
        """
        
        # Create new database connection to avoid concurrency conflicts
        temp_conn = self.get_temp_connection()
        
        try:
            cursor = temp_conn.cursor()
            cursor.execute(courses_sql, subject)
            
            columns = [col[0] for col in cursor.description]
            rows = cursor.fetchall()
            
            courses_data = []
            for row in rows:
                courses_data.append(dict(zip(columns, row)))
            
            return courses_data
        finally:
            temp_conn.close()
    
    def get_classes_by_course_codes(self, course_codes):
        """Get all class sections for selected course codes"""
        if not course_codes:
            return []
        
        placeholders = ','.join(['?' for _ in course_codes])
        # Fix: Simplify query, don't rely on CourseOffering JOIN, because new ClassSection's Offer_Nbr may be NULL
        classes_sql = f"""
        SELECT 
            cs.Class_Nbr, cs.Catalog, cs.Section, cs.Cap_Enrl, cs.Tot_Enrl, cs.Class_Stat,
            cc.Course_Title, cc.Subject, cc.Course_Code, cc.Max_Units,
            COALESCE(co.Term, '2401') as Term, 
            COALESCE(co.Session, 'FAL') as Session, 
            COALESCE(co.Assign_Type, 'CLS') as Assign_Type, 
            COALESCE(co.Component, 'LEC') as Component
        FROM ClassSection cs
        JOIN CourseCatalog cc ON cs.Catalog = cc.Catalog
        LEFT JOIN CourseOffering co ON cs.Catalog = co.Catalog AND 
                  (cs.Offer_Nbr = co.Offer_Nbr OR (cs.Offer_Nbr IS NULL AND co.Offer_Nbr = 1))
        WHERE cc.Course_Code IN ({placeholders}) AND cs.Class_Stat = 'A'
        ORDER BY cc.Course_Code, cs.Section
        """
        
        # Create new database connection to avoid concurrency conflicts
        temp_conn = self.get_temp_connection()
        
        try:
            cursor = temp_conn.cursor()
            cursor.execute(classes_sql, course_codes)
            
            columns = [col[0] for col in cursor.description]
            rows = cursor.fetchall()
            
            classes_data = []
            for row in rows:
                classes_data.append(dict(zip(columns, row)))
            
            return classes_data
        finally:
            temp_conn.close()
    
    def get_teachers_for_classes(self, class_nbrs):
        """Get teachers assigned to specific classes"""
        if not class_nbrs:
            return []
        
        placeholders = ','.join(['?' for _ in class_nbrs])
        teachers_sql = f"""
        SELECT DISTINCT t.F_ID, t.First_Name, t.Last_Name, ci.Class_Nbr
        FROM Teacher t
        JOIN ClassInstructor ci ON t.F_ID = ci.F_ID
        WHERE ci.Class_Nbr IN ({placeholders}) AND ci.Role = 'PI'
        ORDER BY t.Last_Name, t.First_Name
        """
        
        # Create new database connection to avoid concurrency conflicts
        temp_conn = self.get_temp_connection()
        
        try:
            cursor = temp_conn.cursor()
            cursor.execute(teachers_sql, class_nbrs)
            
            columns = [col[0] for col in cursor.description]
            rows = cursor.fetchall()
            
            teachers_data = []
            for row in rows:
                teachers_data.append(dict(zip(columns, row)))
            
            return teachers_data
        finally:
            temp_conn.close()
    
    def get_available_rooms(self, min_capacity=0):
        """Get available rooms with minimum capacity"""
        if 'rooms' not in self.available_options:
            self.load_available_resources()
        
        rooms = self.available_options['rooms']
        filtered_rooms = rooms[rooms['Capacity'] >= min_capacity]
        return filtered_rooms.to_dict('records')
    
    def set_selections(self, **selections):
        """Set user selections"""
        for key, value in selections.items():
            if key in self.selections:
                self.selections[key] = value
                print(f"Selected {key}: {value}")
        
        return self.selections
    
    def generate_timetable(self):
        """Generate timetable based on current selections - 不自动导出Excel"""
        print("Generating timetable...")
        
        # Validate selections
        if not self.selections['classes']:
            return {'error': 'No classes selected'}
        
        # Get selected classes data
        class_placeholders = ','.join(['?' for _ in self.selections['classes']])
        classes_sql = f"""
        SELECT 
            cs.Class_Nbr, cs.Catalog, cs.Section, cs.Cap_Enrl, cs.Tot_Enrl,
            cc.Course_Title, cc.Subject, cc.Course_Code, cc.Max_Units,
            co.Term, co.Session, co.Assign_Type, co.Component
        FROM ClassSection cs
        JOIN CourseCatalog cc ON cs.Catalog = cc.Catalog
        JOIN CourseOffering co ON cs.Catalog = co.Catalog AND cs.Offer_Nbr = co.Offer_Nbr
        WHERE cs.Class_Nbr IN ({class_placeholders}) AND cs.Class_Stat = 'A'
        """
        
        # Use pyodbc connection to avoid SQLAlchemy parameter issues
        cursor = self.conn.cursor()
        cursor.execute(classes_sql, self.selections['classes'])
        
        columns = [col[0] for col in cursor.description]
        rows = cursor.fetchall()
        
        classes_data = []
        for row in rows:
            classes_data.append(dict(zip(columns, row)))
        
        # Generate time slots
        time_slots = self.get_available_time_slots()  # Use available time slots (excluding disabled)
        
        # Get available rooms and teachers
        available_rooms = self.selections['rooms'] if self.selections['rooms'] else [r['Room_ID'] for r in self.get_available_rooms()]
        available_teachers = self.selections['teachers'] if self.selections['teachers'] else []
        
        # Schedule classes
        scheduled_sessions = []
        conflicts = []
        
        # Resource tracking
        teacher_schedule = {}
        room_schedule = {}
        
        for class_info in classes_data:
            class_nbr = class_info['Class_Nbr']
            required_capacity = class_info['Cap_Enrl']
            
            # Get assigned teacher for this class
            teacher_sql = """
            SELECT t.F_ID, t.First_Name, t.Last_Name
            FROM Teacher t
            JOIN ClassInstructor ci ON t.F_ID = ci.F_ID
            WHERE ci.Class_Nbr = ? AND ci.Role = 'PI'
            """
            
            # Use pyodbc to query teacher information
            teacher_cursor = self.conn.cursor()
            teacher_cursor.execute(teacher_sql, class_nbr)
            teacher_rows = teacher_cursor.fetchall()
            
            assigned_teacher = None
            if teacher_rows:
                assigned_teacher = teacher_rows[0][0]  # F_ID
                # If specific teachers selected, check if this teacher is in the list
                if available_teachers and assigned_teacher not in available_teachers:
                    assigned_teacher = None
            
            # Try to schedule this class
            scheduled = self._schedule_single_class(
                class_info, available_rooms, time_slots, 
                teacher_schedule, room_schedule, assigned_teacher
            )
            
            if scheduled:
                scheduled_sessions.append(scheduled)
            else:
                # Create conflict record
                conflict_info = self._create_conflict_record(
                    class_info, available_rooms, assigned_teacher, required_capacity
                )
                conflicts.append(conflict_info)
        
        # Save results to database
        if scheduled_sessions:
            self._save_scheduled_sessions(scheduled_sessions)
        
        if conflicts:
            self._save_conflicts(conflicts)
        
        # Save current schedule results to memory
        self.current_schedule_results = {
            'scheduled_sessions': scheduled_sessions,
            'conflicts': conflicts,
            'generated': True,
            'timestamp': datetime.now()
        }
        
        # Generate timetable view (do not generate Excel file)
        timetable_data = self._generate_timetable_view(scheduled_sessions)
        
        return {
            'success': True,
            'scheduled_count': len(scheduled_sessions),
            'conflict_count': len(conflicts),
            'timetable': timetable_data,
            'conflicts': conflicts,
            'available_time_slots': len(time_slots)
        }
    
    def _generate_time_slots(self):
        """Generate available time slots"""
        time_slots = []
        days = ["Sunday","Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
        periods = [
            ("08:00:00", "09:15:00", "Period 1"),
            ("09:30:00", "10:45:00", "Period 2"),
            ("11:00:00", "12:15:00", "Period 3"),
            ("13:00:00", "14:15:00", "Period 4"),
            ("14:30:00", "15:45:00", "Period 5"),
            ("16:00:00", "17:15:00", "Period 6"),
            ("17:30:00", "18:45:00", "Period 7"),
        ]
        
        for day in days:
            for start, end, name in periods:
                time_slots.append({
                    'day': day,
                    'start_time': start,
                    'end_time': end,
                    'period_name': name,
                    'time_id': f"{day}_{start[:5]}-{end[:5]}"  # Keep time_id format (for disabled time slot matching)
                })
        
        return time_slots
    
    def _schedule_single_class(self, class_info, available_rooms, time_slots, 
                              teacher_schedule, room_schedule, assigned_teacher):
        """Schedule a single class"""
        class_nbr = class_info['Class_Nbr']
        required_capacity = class_info['Cap_Enrl']
        
        # Shuffle time slots for variety
        random.shuffle(time_slots)
        
        for time_slot in time_slots:
            time_id = time_slot['time_id']
            
            # Check teacher availability
            teacher_key = (assigned_teacher, time_id)
            if assigned_teacher and teacher_key in teacher_schedule:
                continue
            
            # Find suitable room
            suitable_rooms = []
            if 'rooms' not in self.available_options:
                self.load_available_resources()
            
            all_rooms = self.available_options['rooms']
            for room_id in available_rooms:
                room_info = all_rooms[all_rooms['Room_ID'] == room_id]
                if len(room_info) > 0 and room_info.iloc[0]['Capacity'] >= required_capacity:
                    suitable_rooms.append(room_info.iloc[0])
            
            for room in suitable_rooms:
                room_id = room['Room_ID']
                room_key = (room_id, time_id)
                
                # Check room availability
                if room_key in room_schedule:
                    continue
                
                # Success! Create session record
                session_record = {
                    'Class_Nbr': class_nbr,
                    'Day': time_slot['day'],
                    'Mtg_Start': time_slot['start_time'],
                    'Mtg_End': time_slot['end_time'],
                    'Room_ID': room_id,
                    'Facil_ID': room.get('Facil_ID', room_id),
                    'Campus': self.selections.get('campus', 'AD'),  # Use user-selected campus
                    'Start_Date': '2024-09-01',
                    'End_Date': '2024-12-15',
                    'F_ID': assigned_teacher,
                    'Room_Capacity': room['Capacity']
                }
                
                # Mark resources as used
                if assigned_teacher:
                    teacher_schedule[teacher_key] = class_nbr
                room_schedule[room_key] = class_nbr
                
                return session_record
        
        return None
    
    def _create_conflict_record(self, class_info, available_rooms, assigned_teacher, required_capacity):
        """Create conflict record with proper field order"""
        return {
            'room_id': None,  # Unable to assign room
            'day': None,      # Unable to assign date
            'Course_ID': class_info.get('Course_Code', ''),
            'Section': class_info['Section'],
            'Mtg_Start': None,
            'Mtg_End': None,
            'Class_Nbr': class_info['Class_Nbr'],
            'Term': class_info['Term'],
            'Session': class_info['Session'],
            'Catalog': class_info['Catalog'],
            'Course_Code': class_info['Course_Code'],
            'Course_Title': class_info['Course_Title'],
            'Subject': class_info['Subject'],
            'Component': class_info.get('Component', ''),
            'F_ID': assigned_teacher,
            'Cap_Enrl': required_capacity,
            'Tot_Enrl': class_info['Tot_Enrl'],
            'Conflict_Type': 'No_Suitable_Time_Room',
            'Conflict_Reason': f'Unable to find suitable time and room for capacity {required_capacity}',
            'Required_Room_Capacity': required_capacity,
            'Campus': self.selections.get('campus', 'AD')  # Use user-selected campus
        }
    
    def _extract_campus_from_location(self, location):
        """Extract campus from room location"""
        if pd.isna(location):
            return 'AD'
        location_str = str(location).upper()
        if 'AD' in location_str:
            return 'AD'
        elif 'AA' in location_str:
            return 'AA'
        elif 'DB' in location_str:
            return 'DB'
        else:
            return 'AD'
    
    def _save_scheduled_sessions(self, scheduled_sessions):
        """Save scheduled sessions to database"""
        cursor = self.conn.cursor()
        
        # Clear existing sessions for selected classes
        class_list = ','.join([str(s['Class_Nbr']) for s in scheduled_sessions])
        cursor.execute(f"DELETE FROM ClassSession WHERE Class_Nbr IN ({class_list})")
        
        # Insert new sessions
        for session in scheduled_sessions:
            sql = """
            INSERT INTO ClassSession 
            (Class_Nbr, Day, Mtg_Start, Mtg_End, Start_Date, End_Date, Room_ID, Facil_ID, Campus)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """
            cursor.execute(sql, (
                session['Class_Nbr'], session['Day'], session['Mtg_Start'], session['Mtg_End'],
                session['Start_Date'], session['End_Date'], session['Room_ID'], 
                session['Facil_ID'], session['Campus']
            ))
        
        self.conn.commit()
        print(f"Saved {len(scheduled_sessions)} scheduled sessions to database")
    
    def _save_conflicts(self, conflicts):
        """Save conflicts to database"""
        cursor = self.conn.cursor()
        
        # Clear existing conflicts for selected classes
        class_list = ','.join([str(c['Class_Nbr']) for c in conflicts])
        cursor.execute(f"DELETE FROM SchedulingConflicts WHERE Class_Nbr IN ({class_list})")
        
        # Insert new conflicts
        for conflict in conflicts:
            sql = """
            INSERT INTO SchedulingConflicts 
            (Class_Nbr, Term, Session, Catalog, Course_Code, Course_Title, Subject, 
             Section, Component, F_ID, Cap_Enrl, Tot_Enrl, 
             Conflict_Type, Conflict_Reason, Required_Room_Capacity, Status)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'PENDING')
            """
            cursor.execute(sql, (
                conflict['Class_Nbr'], conflict['Term'], conflict['Session'], conflict['Catalog'],
                conflict['Course_Code'], conflict['Course_Title'], conflict['Subject'], 
                conflict['Section'], conflict['Component'], conflict['F_ID'], 
                conflict['Cap_Enrl'], conflict['Tot_Enrl'],
                conflict['Conflict_Type'], conflict['Conflict_Reason'], conflict['Required_Room_Capacity']
            ))
        
        self.conn.commit()
        print(f"Saved {len(conflicts)} conflicts to database")
    
    def _generate_timetable_view(self, scheduled_sessions):
        """Generate timetable view for web display"""
        if not scheduled_sessions:
            return {}
        
        # Organize by day and time
        timetable = {}
        days = ["Sunday","Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
        times = ["08:00:00-09:15:00", "09:30:00-10:45:00", "11:00:00-12:15:00", "13:00:00-14:15:00", 
                "14:30:00-15:45:00", "16:00:00-17:15:00", "17:30:00-18:45:00"]
        
        for day in days:
            timetable[day] = {}
            for time in times:
                timetable[day][time] = []
        
        # Get class details for each session
        for session in scheduled_sessions:
            class_nbr = session['Class_Nbr']
            day = session['Day']
            time_slot = f"{session['Mtg_Start']}-{session['Mtg_End']}"
            
            # Get class details
            class_sql = """
            SELECT 
                cs.Class_Nbr, cs.Section, cc.Course_Code, cc.Course_Title, cc.Subject,
                t.First_Name, t.Last_Name, r.Description as Room_Description
            FROM ClassSection cs
            JOIN CourseCatalog cc ON cs.Catalog = cc.Catalog
            LEFT JOIN ClassInstructor ci ON cs.Class_Nbr = ci.Class_Nbr AND ci.Role = 'PI'
            LEFT JOIN Teacher t ON ci.F_ID = t.F_ID
            LEFT JOIN Room r ON ? = r.Room_ID
            WHERE cs.Class_Nbr = ?
            """
            
            # Use pyodbc to query course details
            details_cursor = self.conn.cursor()
            details_cursor.execute(class_sql, (session['Room_ID'], class_nbr))
            details_rows = details_cursor.fetchall()
            
            if details_rows:
                details_row = details_rows[0]
                if day in timetable and time_slot in timetable[day]:
                    timetable[day][time_slot].append({
                        'class_nbr': class_nbr,
                        'course_code': details_row[2],  # Course_Code
                        'course_title': details_row[3],  # Course_Title
                        'section': details_row[1],  # Section
                        'subject': details_row[4],  # Subject
                        'teacher': f"{details_row[5]} {details_row[6]}" if details_row[5] else 'TBD',  # First_Name Last_Name
                        'room': session['Room_ID'],
                        'room_description': details_row[7] if details_row[7] else session['Room_ID']  # Room_Description
                    })
        
        return timetable
    
    
    
    
    def export_schedule_results_to_excel(self):
        """Export the latest schedule results including scheduled sessions and conflicts"""
        try:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            filename = f"Schedule_Results_{timestamp}.xlsx"
            
            # Check if schedule results have been generated
            if not self.current_schedule_results['generated']:
                return {
                    'success': False,
                    'error': 'No schedule results available. Please generate a timetable first.'
                }
            
            scheduled_sessions = self.current_schedule_results['scheduled_sessions']
            conflicts = self.current_schedule_results['conflicts']
            
            if not scheduled_sessions and not conflicts:
                return {
                    'success': False,
                    'error': 'No schedule data to export.'
                }
            
            # Get complete scheduled course information
            scheduled_data = []
            if scheduled_sessions:
                class_nbrs = [str(session['Class_Nbr']) for session in scheduled_sessions]
                if class_nbrs:
                    # Create a mapping from class_nbr to session data
                    session_map = {str(session['Class_Nbr']): session for session in scheduled_sessions}
                    
                    # Use a SQL query to get complete information for all courses
                    placeholders = ','.join(['?' for _ in class_nbrs])
                    complete_sql = f"""
                    SELECT 
                        co.Access, co.Term, co.Assign_Type, cs.Class_Nbr, co.Offer_Nbr, cc.Max_Units,
                        cs.Enrl_Stat, cc.Long_Title, co.Component, cs.Catalog, cc.Acad_Group,
                        NULL as Pat, NULL as Pat_Nbr, co.Session, ci.F_ID, t.First_Name, t.Last_Name,
                        ci.Role, cc.Career, term.Start_Date, term.End_Date, cc.Course_ID, cc.Course_Code,
                        cc.Subject, cc.Descr, cs.Section, cs.Class_Stat, 
                        NULL as Mtg_Start, NULL as Mtg_End, NULL as Campus, 
                        cs.Tot_Enrl, cs.Cap_Enrl, NULL as Facil_ID, NULL as Day, NULL as Room_ID,
                        NULL as Room_Capacity
                    FROM ClassSection cs
                    JOIN CourseCatalog cc ON cs.Catalog = cc.Catalog
                    JOIN CourseOffering co ON cs.Catalog = co.Catalog AND cs.Offer_Nbr = co.Offer_Nbr
                    LEFT JOIN Term term ON co.Term = term.Term_Code AND co.Session = term.Session
                    LEFT JOIN ClassInstructor ci ON cs.Class_Nbr = ci.Class_Nbr AND ci.Role = 'PI'
                    LEFT JOIN Teacher t ON ci.F_ID = t.F_ID
                    WHERE cs.Class_Nbr IN ({placeholders})
                    """
                    
                    try:
                        cursor = self.conn.cursor()
                        cursor.execute(complete_sql, class_nbrs)
                        
                        for row in cursor.fetchall():
                            # Create a record with all standard columns
                            record = {}
                            for i, col_name in enumerate(self.standard_columns):
                                record[col_name] = row[i] if i < len(row) else None
                            
                            # Get scheduled session information from session_map and overwrite
                            class_nbr = str(row[3])  # Class_Nbr is in the 4th position
                            if class_nbr in session_map:
                                session_data = session_map[class_nbr]
                                
                                record.update({
                                    'Mtg_Start': session_data.get('Mtg_Start'),
                                    'Mtg_End': session_data.get('Mtg_End'),
                                    'Campus': session_data.get('Campus'),
                                    'Day': session_data.get('Day'),
                                    'Room_ID': session_data.get('Room_ID'),
                                    'Facil_ID': session_data.get('Facil_ID'),
                                    'Room_Capacity': session_data.get('Room_Capacity'),
                                    'Start_Date': session_data.get('Start_Date'),
                                    'End_Date': session_data.get('End_Date')
                                })
                            
                            scheduled_data.append(record)
                    except Exception as e:
                        print(f"Error getting complete course information: {str(e)}")
                        # If query fails, create basic record
                        for session in scheduled_sessions:
                            basic_record = {col: None for col in self.standard_columns}
                            basic_record.update({
                                'Class_Nbr': session.get('Class_Nbr'),
                                'Day': session.get('Day'),
                                'Mtg_Start': session.get('Mtg_Start'),
                                'Mtg_End': session.get('Mtg_End'),
                                'Room_ID': session.get('Room_ID'),
                                'Campus': session.get('Campus'),
                                'Start_Date': session.get('Start_Date'),
                                'End_Date': session.get('End_Date'),
                                'F_ID': session.get('F_ID'),
                                'Room_Capacity': session.get('Room_Capacity')
                            })
                            scheduled_data.append(basic_record)
            
            # Create DataFrame, ensure standard column order
            scheduled_df = pd.DataFrame(scheduled_data)
            if len(scheduled_df) > 0:
                scheduled_df = scheduled_df.reindex(columns=self.standard_columns, fill_value=None)
            
            # Save to Excel - only export scheduled sessions, not conflicts
            with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                if len(scheduled_df) > 0:
                    scheduled_df.to_excel(writer, sheet_name='Scheduled_Sessions', index=False)
                
                # Add summary information page
                summary_data = {
                    'Item': ['Generation Time', 'Scheduled Sessions', 'Conflicts', 'Total Classes'],
                    'Value': [
                        self.current_schedule_results['timestamp'].strftime('%Y-%m-%d %H:%M:%S'),
                        len(scheduled_sessions),
                        len(conflicts),
                        len(scheduled_sessions) + len(conflicts)
                    ]
                }
                summary_df = pd.DataFrame(summary_data)
                summary_df.to_excel(writer, sheet_name='Summary', index=False)
            
            print(f"Excel file generated: {filename}")
            print(f"Contains complete {len(self.standard_columns)} columns")
            print(f"Only export scheduled courses, not conflicts")
            return {
                'success': True,
                'filename': filename,
                'record_count': len(scheduled_df),  # Only calculate scheduled sessions
                'scheduled_count': len(scheduled_df),
                'conflicts_count': len(conflicts)  # Still return conflict count for statistics, but not exported
            }
                
        except Exception as e:
            print(f"Error generating Excel file: {str(e)}")
            return {
                'success': False,
                'error': str(e)
            }
    
    def import_excel_data(self, file_path):
        """Simulate Excel data import (for testing purposes)"""
        try:
            # Read Excel file for simulation
            df = pd.read_excel(file_path)
            
            # 模拟导入统计
            simulated_data = {
                'Faculty': 15,
                'Room': 25, 
                'Course': 120,
                'Campus': 3,
                'Students': 450,
                'Class_Sections': 180
            }
            
            return {
                'success': True,
                'message': 'Successfully imported (simulated)',
                'imported_data': simulated_data,
                'total_rows': len(df),
                'file_info': f"File: {file_path}, Columns: {len(df.columns)}"
            }
        except Exception as e:
            return {
                'success': False,
                'error': str(e)
            }
    
    def export_all_data_to_excel(self):
        """Export all scheduling data to Excel file"""
        try:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            filename = f"Full_Schedule_Export_{timestamp}.xlsx"
            
            # Get all scheduled sessions with full details
            export_sql = """
            SELECT 
                co.Access, co.Term, co.Assign_Type, sess.Class_Nbr, co.Offer_Nbr, cc.Max_Units,
                cls.Enrl_Stat, cc.Long_Title, co.Component, cls.Catalog, cc.Acad_Group,
                NULL as Pat, NULL as Pat_Nbr, co.Session, ci.F_ID, t.First_Name, t.Last_Name,
                ci.Role, cc.Career, sess.Start_Date, sess.End_Date, cc.Course_ID, cc.Course_Code,
                cc.Subject, cc.Descr, cls.Section, cls.Class_Stat, sess.Mtg_Start, sess.Mtg_End,
                sess.Campus, cls.Tot_Enrl, cls.Cap_Enrl, sess.Facil_ID, sess.Day, sess.Room_ID,
                r.Capacity as Room_Capacity
            FROM ClassSession sess
            JOIN ClassSection cls ON sess.Class_Nbr = cls.Class_Nbr
            JOIN CourseCatalog cc ON cls.Catalog = cc.Catalog
            JOIN CourseOffering co ON cls.Catalog = co.Catalog AND cls.Offer_Nbr = co.Offer_Nbr
            LEFT JOIN ClassInstructor ci ON sess.Class_Nbr = ci.Class_Nbr AND ci.Role = 'PI'
            LEFT JOIN Teacher t ON ci.F_ID = t.F_ID
            LEFT JOIN Room r ON sess.Room_ID = r.Room_ID
            ORDER BY sess.Day, sess.Mtg_Start
            """
            
            # Use pyodbc to execute query
            cursor = self.conn.cursor()
            cursor.execute(export_sql)
            
            columns = [col[0] for col in cursor.description]
            rows = cursor.fetchall()
            
            # Convert to DataFrame
            export_data = []
            for row in rows:
                export_data.append(dict(zip(columns, row)))
            
            df = pd.DataFrame(export_data)
            
            # Reorder columns to match standard format
            df = df.reindex(columns=self.standard_columns, fill_value=None)
            
            
            # Save to Excel
            with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Schedule_Data', index=False)
            
            return {
                'success': True,
                'filename': filename,
                'record_count': len(df)
            }
        except Exception as e:
            return {
                'success': False,
                'error': str(e)
            }

    def get_available_time_slots(self):
        """Get available time slots (excluding disabled ones)"""
        all_slots = self._generate_time_slots()
        available_slots = [slot for slot in all_slots if slot['time_id'] not in self.disabled_time_slots]
        return available_slots
    
    def get_schedule_results_status(self):
        """ Get current schedule results status"""
        return {
            'has_results': self.current_schedule_results['generated'],
            'scheduled_count': len(self.current_schedule_results['scheduled_sessions']),
            'conflicts_count': len(self.current_schedule_results['conflicts']),
            'timestamp': self.current_schedule_results['timestamp'].isoformat() if self.current_schedule_results['timestamp'] else None
        }
    
    def disable_time_slots(self, time_slot_patterns):
        """Batch disable time slots - now accepts exact time slot ID list"""
        # Clear existing disabled time slots
        self.disabled_time_slots.clear()
        
        # Add new disabled time slots
        for time_slot_id in time_slot_patterns:
            if isinstance(time_slot_id, str) and time_slot_id.strip():
                self.disabled_time_slots.add(time_slot_id.strip())
        
        return {
            'success': True,
            'disabled_count': len(self.disabled_time_slots),
            'disabled_slots': list(self.disabled_time_slots)
        }
    
    def enable_all_time_slots(self):
        """Re-enable all time slots"""
        self.disabled_time_slots.clear()
        return {'success': True, 'message': 'All time slots enabled'}

    def get_smart_rooms_for_courses(self, course_codes):
        """Smartly recommend suitable rooms for selected courses"""
        if not course_codes:
            return self.get_available_rooms()
        
        # Get course information
        placeholders = ','.join(['?' for _ in course_codes])
        courses_sql = f"""
        SELECT DISTINCT Course_Code, Course_Title, Subject
        FROM CourseCatalog 
        WHERE Course_Code IN ({placeholders})
        """
        
        cursor = self.conn.cursor()
        cursor.execute(courses_sql, course_codes)
        courses = cursor.fetchall()
        
        # Analyze course types
        lab_keywords = ['lab', 'laboratory', 'practical', 'experiment', 'workshop']
        computer_keywords = ['computer', 'programming', 'software', 'coding', 'it', 'csc', 'cse', 'ite']
        engineering_keywords = ['engineering', 'mec', 'civ', 'eee', 'bme', 'cad']
        
        needs_lab = False
        needs_computer_lab = False
        needs_engineering_lab = False
        needs_special_facility = False
        
        for course in courses:
            course_code, course_title, subject = course
            title_lower = str(course_title).lower() if course_title else ""
            code_lower = str(course_code).lower() if course_code else ""
            subject_lower = str(subject).lower() if subject else ""
            
            # Check if lab is needed
            if any(keyword in title_lower for keyword in lab_keywords):
                needs_lab = True
                
            # Check if computer lab is needed
            if any(keyword in title_lower or keyword in code_lower or keyword in subject_lower 
                   for keyword in computer_keywords):
                needs_computer_lab = True
                
            # Check if engineering lab is needed
            if any(keyword in title_lower or keyword in subject_lower 
                   for keyword in engineering_keywords):
                needs_engineering_lab = True
                
            # Check for special facility needs
            if ('aviation' in title_lower or 'avs' in subject_lower or
                'architecture' in title_lower or 'arc' in subject_lower or
                'chemistry' in title_lower or 'bio' in title_lower):
                needs_special_facility = True
        
        # Build Room filter conditions
        room_filter_conditions = []
        
        if needs_lab or needs_computer_lab or needs_engineering_lab or needs_special_facility:
            # Need special facilities, prioritize matching labs
            lab_conditions = []
            
            if needs_computer_lab:
                lab_conditions.append("Description LIKE '%Computer Lab%'")
                lab_conditions.append("Description LIKE '%CAD Lab%'")
                
            if needs_engineering_lab:
                lab_conditions.append("Description LIKE '%Engineering%'")
                lab_conditions.append("Description LIKE '%CAD Lab%'")
                lab_conditions.append("Description LIKE '%Circuit Lab%'")
                lab_conditions.append("Description LIKE '%Control%Lab%'")
                
            if needs_special_facility:
                lab_conditions.append("Description LIKE '%Aviation Lab%'")
                lab_conditions.append("Description LIKE '%Bio Lab%'")
                lab_conditions.append("Description LIKE '%Chemistry Lab%'")
                lab_conditions.append("Description LIKE '%Arch%'")
                
            if needs_lab and not lab_conditions:
                # General lab needs
                lab_conditions.append("Description LIKE '%Lab%'")
                lab_conditions.append("Description LIKE '%Laboratory%'")
            
            if lab_conditions:
                room_filter_conditions.append(f"({' OR '.join(lab_conditions)})")
        
        # Always include regular classroom as backup
        room_filter_conditions.append("""(Description LIKE '%Classroom%' 
                                          OR Description LIKE '%Class%'
                                          OR Description = 'Classroom')""")
        
        # Build final query
        where_clause = f"WHERE ({' OR '.join(room_filter_conditions)})" if room_filter_conditions else ""
        
        rooms_sql = f"""
        SELECT Room_ID, Description, Capacity, Gender, Location, Facil_ID
        FROM Room 
        {where_clause}
        AND Capacity > 0
        ORDER BY 
            CASE 
                WHEN Description LIKE '%Lab%' THEN 1
                WHEN Description LIKE '%Classroom%' THEN 2
                ELSE 3
            END,
            Capacity DESC
        """
        
        cursor.execute(rooms_sql)
        columns = [col[0] for col in cursor.description]
        rows = cursor.fetchall()
        
        rooms_data = []
        for row in rows:
            rooms_data.append(dict(zip(columns, row)))
        
        return {
            'rooms': rooms_data,
            'match_info': {
                'needs_lab': needs_lab,
                'needs_computer_lab': needs_computer_lab,
                'needs_engineering_lab': needs_engineering_lab,
                'needs_special_facility': needs_special_facility,
                'total_rooms': len(rooms_data)
            }
        }

    def get_subjects_by_multiple_acad_groups(self, acad_groups):
        """Get subjects from multiple academic groups"""
        if not acad_groups or len(acad_groups) == 0:
            return {'regular_subjects': []}
        
        # Convert single string to list if needed
        if isinstance(acad_groups, str):
            acad_groups = [acad_groups]
        
        # Get subjects from selected academic groups
        placeholders = ','.join(['?' for _ in acad_groups])
        subject_sql = f"""
        SELECT cc.Subject, cc.Acad_Group, 
               COUNT(DISTINCT cc.Course_Code) as Course_Count
        FROM CourseCatalog cc 
        WHERE cc.Acad_Group IN ({placeholders}) 
          AND cc.Subject IS NOT NULL 
          AND cc.Course_Code IS NOT NULL AND cc.Course_Code != ''
        GROUP BY cc.Subject, cc.Acad_Group 
        ORDER BY cc.Acad_Group, cc.Subject
        """
        
        # Create new database connection to avoid concurrency conflicts
        temp_conn = self.get_temp_connection()
        
        try:
            cursor = temp_conn.cursor()
            cursor.execute(subject_sql, acad_groups)
            
            columns = [col[0] for col in cursor.description]
            rows = cursor.fetchall()
            
            regular_subjects = []
            for row in rows:
                regular_subjects.append(dict(zip(columns, row)))
            
            return {
                'regular_subjects': regular_subjects,
                'selected_groups': acad_groups,
                'is_multi_group': len(acad_groups) > 1
            }
        finally:
            temp_conn.close()
    
    def get_cross_departmental_subjects(self, acad_groups):
        """Get subjects that appear in multiple selected academic groups"""
        if not acad_groups or len(acad_groups) < 2:
            return []
        
        placeholders = ','.join(['?' for _ in acad_groups])
        cross_sql = f"""
        SELECT cc.Subject, 
               COUNT(DISTINCT cc.Acad_Group) as Group_Count,
               COUNT(DISTINCT cc.Course_Code) as Course_Count
        FROM CourseCatalog cc 
        WHERE cc.Acad_Group IN ({placeholders})
          AND cc.Subject IS NOT NULL 
          AND cc.Course_Code IS NOT NULL AND cc.Course_Code != ''
        GROUP BY cc.Subject
        HAVING COUNT(DISTINCT cc.Acad_Group) > 1
        ORDER BY COUNT(DISTINCT cc.Acad_Group) DESC, cc.Subject
        """
        
        # Create new database connection to avoid concurrency conflicts
        temp_conn = self.get_temp_connection()
        
        try:
            cursor = temp_conn.cursor()
            cursor.execute(cross_sql, acad_groups)
            
            cross_subjects = []
            for row in cursor.fetchall():
                subject, group_count, course_count = row
                
                # Get detailed groups for this subject
                group_sql = f"""
                SELECT DISTINCT Acad_Group 
                FROM CourseCatalog 
                WHERE Subject = ? AND Acad_Group IN ({placeholders})
                ORDER BY Acad_Group
                """
                cursor.execute(group_sql, [subject] + acad_groups)
                groups = [g[0] for g in cursor.fetchall()]
                
                cross_subjects.append({
                    'Subject': subject,
                    'Course_Count': course_count,
                    'Group_Count': group_count,
                    'Academic_Groups': groups
                })
            
            return cross_subjects
        finally:
            temp_conn.close()

# Flask Web API Interface
def create_web_api():
    """Create Flask web API for the scheduling system"""
    from flask import Flask, request, jsonify, render_template_string
    
    app = Flask(__name__)
    scheduler = WebSchedulingSystem()
    
    # HTML Template for the web interface
    HTML_TEMPLATE = """
    <!DOCTYPE html>
    <html>
    <head>
        <title>Hierarchical Scheduling System</title>
        <style>
            body { font-family: Arial, sans-serif; margin: 20px; }
            .container { max-width: 1400px; margin: 0 auto; }
            .section { margin: 20px 0; padding: 15px; border: 1px solid #ddd; border-radius: 5px; }
            .selection-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); gap: 20px; }
            select, input, button { margin: 5px; padding: 8px; }
            button { background-color: #007bff; color: white; border: none; border-radius: 3px; cursor: pointer; }
            button:hover { background-color: #0056b3; }
            .timetable { width: 100%; border-collapse: collapse; margin: 20px 0; }
            .timetable th, .timetable td { border: 1px solid #ddd; padding: 8px; text-align: center; }
            .timetable th { background-color: #f2f2f2; }
            .class-item { background-color: #e3f2fd; margin: 2px; padding: 4px; border-radius: 3px; font-size: 12px; }
            .results { margin: 20px 0; padding: 15px; background-color: #f8f9fa; border-radius: 5px; }
            
            /* Time slot management style */
            .timeslot-table { width: 100%; border-collapse: collapse; margin: 20px 0; }
            .timeslot-table th, .timeslot-table td { border: 1px solid #ddd; padding: 10px; text-align: center; }
            .timeslot-table th { background-color: #f2f2f2; font-weight: bold; }
            .timeslot-cell { cursor: pointer; transition: background-color 0.3s; }
            .timeslot-available { background-color: #d4edda; color: #155724; }
            .timeslot-disabled { background-color: #f8d7da; color: #721c24; }
            .timeslot-cell:hover { opacity: 0.8; }
            
            /* Selection box style */
            .selection-box { max-height: 200px; overflow-y: auto; border: 1px solid #ddd; padding: 10px; border-radius: 5px; }
            .checkbox-item { margin: 5px 0; }
            .checkbox-item label { display: flex; align-items: center; cursor: pointer; }
            .checkbox-item input[type="checkbox"] { margin-right: 8px; }
            
            /* Button group style */
            .button-group { margin: 15px 0; }
            .button-group button { margin-right: 10px; }
            
            /* Cross-departmental subject style */
            .cross-dept-subject { background-color: #e8f5e8; margin: 5px 0; padding: 10px; border-radius: 5px; border-left: 4px solid #28a745; }
            .cross-dept-subject h5 { margin: 0 0 5px 0; color: #155724; }
            .cross-dept-subject .subject-info { font-size: 13px; color: #666; }
            .cross-dept-checkbox { margin: 8px 0; }
            .cross-dept-checkbox input[type="checkbox"] { margin-right: 8px; transform: scale(1.1); }
            
            /* 下载按钮样式 */
            #downloadBtn {
                background: linear-gradient(45deg, #007bff, #0056b3);
                transition: all 0.3s ease;
                box-shadow: 0 2px 4px rgba(0,123,255,0.3);
            }
            #downloadBtn:hover {
                background: linear-gradient(45deg, #0056b3, #004085);
                box-shadow: 0 4px 8px rgba(0,123,255,0.5);
                transform: translateY(-1px);
            }
            #downloadBtn:active {
                transform: translateY(0);
                box-shadow: 0 2px 4px rgba(0,123,255,0.3);
            }
            
            /* File information style */
            #lastGeneratedFile {
                background-color: #f8f9fa;
                border: 1px solid #dee2e6;
                border-radius: 5px;
                padding: 10px;
            }
            #fileName {
                font-weight: bold;
                color: #495057;
            }
            
            .status-legend { display: flex; align-items: center; gap: 20px; margin: 10px 0; }
            .legend-item { display: flex; align-items: center; gap: 5px; }
            .legend-color { width: 20px; height: 20px; border: 1px solid #ddd; }
        </style>
        <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    </head>
    <body>
        <div class="container">
            <h1>Hierarchical Scheduling System</h1>
            
            <div class="selection-grid">
                <!-- Resource Selection Sections -->
                <div class="section">
                    <h3>1. Select Term</h3>
                    <select id="termSelect" onchange="loadSessions()">
                        <option value="">Choose Term...</option>
                    </select>
                </div>
                
                <div class="section">
                    <h3>2. Select Session</h3>
                    <select id="sessionSelect">
                        <option value="">Choose Session...</option>
                    </select>
                </div>
                
                <div class="section">
                    <h3>3. Select Campus</h3>
                    <select id="campusSelect">
                        <option value="">Choose Campus...</option>
                    </select>
                </div>
                
                <div class="section">
                    <h3>4. Select Academic Groups</h3>
                    <div class="selection-box" id="acadGroupSelection">
                        <p>Loading academic groups...</p>
                    </div>
                    <div class="button-group">
                        <button onclick="clearAcadGroupSelection()">Clear Selection</button>
                    </div>
                </div>
                
                <div class="section">
                    <h3>5.1 Select Subject</h3>
                    <select id="subjectSelect" onchange="loadCourses()" disabled>
                        <option value="">Choose Subject...</option>
                    </select>
                    <p id="subjectSelectHint" style="color: #666; font-size: 14px; margin-top: 5px;">Please select a single academic group to enable subject selection</p>
                </div>
                
                <!-- Cross-Departmental Subjects Section -->
                <div class="section" id="crossDepartmentalSection" style="display: none;">
                    <h3>5.2 Cross-Departmental Subjects</h3>
                    <div id="crossDepartmentalContent">
                        <p>Select multiple academic groups to see cross-departmental subjects...</p>
                    </div>
                </div>
                
                <div class="section">
                    <h3 id="courseSelectionTitle">6. Select Courses</h3>
                    <div class="selection-box" id="courseSelection">
                        <p>Please select a subject first...</p>
                    </div>
                </div>
            </div>
            
            <div class="selection-grid">
                <div class="section">
                    <h3 id="teacherSelectionTitle">7. Select Teachers (Optional)</h3>
                    <div class="selection-box" id="teacherSelection">
                        <p>No teachers loaded...</p>
                    </div>
                </div>
                
                <div class="section">
                    <h3 id="roomSelectionTitle">8. Select Rooms (Optional)</h3>
                    <div class="selection-box" id="roomSelection">
                        <p>Loading rooms...</p>
                    </div>
                </div>
            </div>
            
            <!-- Time slot management section -->
            <div class="section">
                <h3>9.Time Slot Management</h3>
                <div class="status-legend">
                    <div class="legend-item">
                        <div class="legend-color timeslot-available"></div>
                        <span>Available</span>
                    </div>
                    <div class="legend-item">
                        <div class="legend-color timeslot-disabled"></div>
                        <span>Disabled</span>
                    </div>
                </div>
                
                <div class="button-group">
                    <button onclick="loadTimeSlotStatus()">Refresh Status</button>
                    <button onclick="enableAllTimeSlots()">Enable All Slots</button>
                    <button onclick="saveTimeSlotSettings()">Save Settings</button>
                </div>
                
                <table class="timeslot-table" id="timeslotTable">
                    <thead>
                        <tr>
                            <th>Time / Day</th>
                            <th>Sunday</th>
                            <th>Monday</th>
                            <th>Tuesday</th>
                            <th>Wednesday</th>
                            <th>Thursday</th>
                            <th>Friday</th>
                            <th>Saturday</th>
                        </tr>
                    </thead>
                    <tbody>
                        <!-- Time slot content will be generated dynamically through JavaScript -->
                    </tbody>
                </table>
            </div>
            
            <div class="section">
                <button onclick="generateSchedule()" style="font-size: 18px; padding: 15px 30px;">
                    Generate Timetable
                </button>
                
                <div style="margin-top: 20px;">
                    <h4>File Operations</h4>
                    <p style="color: #666; font-size: 14px; margin-bottom: 15px;">
                        📋 Export will include all 36 standard columns (Access, Term, Assign_Type, Class_Nbr, etc.)
                    </p>
                    <input type="file" id="excelFileInput" accept=".xlsx,.xls" style="margin: 10px 0;">
                    <button onclick="simulateImport()" style="margin: 5px;">Simulate Import Data</button>
                    <button id="exportScheduleBtn" onclick="exportScheduleResults()" style="margin: 5px; background-color: #28a745;" disabled>Export Schedule Results</button>
                    <div id="lastGeneratedFile" style="margin-top: 10px; display: none;">
                        <span id="fileName"></span>
                        <button id="downloadBtn" onclick="downloadFile()" style="margin-left: 10px; background-color: #007bff; color: white; border: none; padding: 5px 10px; border-radius: 3px;">
                            📥 Download File
                        </button>
                    </div>
                </div>
            </div>
            
            <div id="results" class="results" style="display: none;"></div>
            <div id="timetable"></div>
        </div>
        
        <script>
            // Global variables
            let timeSlotStatus = {};
            let currentDisabledSlots = new Set();
            let lastGeneratedFileName = '';  // Track latest generated file name
            
            // Update selection count function
            function updateSelectionCounts() {
                // Update course selection count
                const selectedCourses = $('input[name="courses"]:checked').length;
                $('#courseSelectionTitle').text(`6. Select Courses${selectedCourses > 0 ? ` (${selectedCourses} selected)` : ''}`);
                
                // Update teacher selection count
                const selectedTeachers = $('input[name="teachers"]:checked').length;
                $('#teacherSelectionTitle').text(`7. Select Teachers (Optional)${selectedTeachers > 0 ? ` (${selectedTeachers} selected)` : ''}`);
                
                // Update room selection count
                const selectedRooms = $('input[name="rooms"]:checked').length;
                $('#roomSelectionTitle').text(`8. Select Rooms (Optional)${selectedRooms > 0 ? ` (${selectedRooms} selected)` : ''}`);
            }
            
            // Bind selection change events
            function bindSelectionChangeEvents() {
                // Course selection change
                $(document).on('change', 'input[name="courses"]', function() {
                    updateSelectionCounts();
                    updateClassSections();
                });
                
                // Teacher selection change
                $(document).on('change', 'input[name="teachers"]', function() {
                    updateSelectionCounts();
                });
                
                // Room selection change
                $(document).on('change', 'input[name="rooms"]', function() {
                    updateSelectionCounts();
                });
            }
            
            // Load initial data
            $(document).ready(function() {
                loadTerms();
                loadCampuses();
                loadAcadGroups();
                loadRooms();
                loadTimeSlotStatus();
                bindSelectionChangeEvents(); // Bind selection change events
                updateSelectionCounts(); // Initialize count display
            });
            
            function loadTerms() {
                $.get('/api/terms', function(data) {
                    const select = $('#termSelect');
                    select.empty().append('<option value="">Choose Term...</option>');
                    data.forEach(term => {
                        select.append(`<option value="${term.Term_Code}">${term.Term_Name} (${term.Session})</option>`);
                    });
                }).fail(function() {
                    console.error('Failed to load terms');
                });
            }
            
            function loadSessions() {
                $.get('/api/sessions', function(data) {
                    const select = $('#sessionSelect');
                    select.empty().append('<option value="">Choose Session...</option>');
                    data.forEach(session => {
                        select.append(`<option value="${session.Session}">${session.Session}</option>`);
                    });
                }).fail(function() {
                    console.error('Failed to load sessions');
                });
            }
            
            function loadCampuses() {
                $.get('/api/campuses', function(data) {
                    const select = $('#campusSelect');
                    select.empty().append('<option value="">Choose Campus...</option>');
                    data.forEach(campus => {
                        select.append(`<option value="${campus.Campus}">${campus.Campus}: ${campus.Description}</option>`);
                    });
                }).fail(function() {
                    console.error('Failed to load campuses');
                });
            }
            
            function loadAcadGroups() {
                $.get('/api/acad_groups', function(data) {
                    const container = $('#acadGroupSelection');
                    container.empty();
                    
                    if (data.length === 0) {
                        container.html('<p>No academic groups available.</p>');
                        return;
                    }
                    
                    container.append('<p><strong>Select Academic Groups (Multi-select supported):</strong></p>');
                    
                    data.forEach(group => {
                        container.append(`
                            <div class="checkbox-item">
                                <label>
                                    <input type="checkbox" name="acadGroups" value="${group.Acad_Group}" onchange="onAcadGroupSelectionChange()">
                                    ${group.Acad_Group}
                                </label>
                            </div>
                        `);
                    });
                }).fail(function() {
                    $('#acadGroupSelection').html('<p style="color: red;">Failed to load academic groups.</p>');
                    console.error('Failed to load academic groups');
                });
            }
            
            function onAcadGroupSelectionChange() {
                const selectedGroups = $('input[name="acadGroups"]:checked').map(function() { 
                    return this.value; 
                }).get();
                
                // Clear previous selections
                $('#subjectSelect').empty().append('<option value="">Choose Subject...</option>');
                $('#crossDepartmentalSection').hide();
                $('#courseSelection').html('<p>Please select a subject first...</p>');
                $('#teacherSelection').html('<p>No courses selected...</p>');
                $('#roomSelection').html('<p>Loading rooms...</p>');
                loadRooms(); // Reload default rooms
                updateSelectionCounts(); // Update count
                
                if (selectedGroups.length === 0) {
                    // No academic groups selected
                    $('#subjectSelect').prop('disabled', true);
                    $('#subjectSelectHint').text('Please select a single academic group to enable subject selection').show();
                    $('#crossDepartmentalSection').hide();
                } else if (selectedGroups.length === 1) {
                    // Selected single academic group - enable 5.1 Select Subject
                    $('#subjectSelect').prop('disabled', false);
                    $('#subjectSelectHint').text('Single academic group selected - you can now select subjects').show();
                    $('#crossDepartmentalSection').hide();
                    
                    // Automatically load subjects for the selected academic group
                    loadSubjectsForSingleGroup(selectedGroups[0]);
                } else {
                    // Selected multiple academic groups - enable 5.2 Cross-Departmental Subjects
                    $('#subjectSelect').prop('disabled', true);
                    $('#subjectSelectHint').text('Multiple academic groups selected - use Cross-Departmental Subjects section below').show();
                    $('#crossDepartmentalSection').show();
                    
                    // Automatically load cross-departmental subjects and regular subjects
                    loadSubjectsForMultipleGroups(selectedGroups);
                }
            }
            
            function loadSubjectsForSingleGroup(acadGroup) {
                $.get(`/api/subjects/${acadGroup}`, function(data) {
                    const select = $('#subjectSelect');
                    select.empty().append('<option value="">Choose Subject...</option>');
                    
                    if (data && data.length > 0) {
                        data.forEach(subject => {
                            select.append(`<option value="${subject.Subject}">${subject.Subject} (${subject.Course_Count} courses)</option>`);
                        });
                    }
                }).fail(function() {
                    console.error('Failed to load subjects for single group');
                    $('#subjectSelect').empty().append('<option value="">Failed to load subjects</option>');
                });
            }
            
            function loadSubjectsForMultipleGroups(selectedGroups) {
                // Load regular subjects to 5.1 section (but keep disabled)
                $.ajax({
                    url: '/api/subjects_multi_groups',
                    type: 'POST',
                    contentType: 'application/json',
                    data: JSON.stringify({acad_groups: selectedGroups}),
                    success: function(data) {
                        updateSubjectSelection(data, selectedGroups);
                    },
                    error: function() {
                        console.error('Failed to load subjects from multiple groups');
                    }
                });
                
                // 加载跨部门学科到5.2部分
                loadCrossDepartmentalSubjects(selectedGroups);
            }
            
            function updateSubjectSelection(data, selectedGroups) {
                    const select = $('#subjectSelect');
                    select.empty().append('<option value="">Choose Subject...</option>');
                
                if (data.regular_subjects && data.regular_subjects.length > 0) {
                    // Group subjects by academic group
                    const groupedSubjects = {};
                    data.regular_subjects.forEach(subject => {
                        if (!groupedSubjects[subject.Acad_Group]) {
                            groupedSubjects[subject.Acad_Group] = [];
                        }
                        groupedSubjects[subject.Acad_Group].push(subject);
                    });
                    
                    // Add grouped options
                    Object.keys(groupedSubjects).sort().forEach(group => {
                        select.append(`<optgroup label="${group}">`);
                        groupedSubjects[group].forEach(subject => {
                        select.append(`<option value="${subject.Subject}">${subject.Subject} (${subject.Course_Count} courses)</option>`);
                    });
                        select.append('</optgroup>');
                    });
                }
            }
            
            function loadCrossDepartmentalSubjects(selectedGroups) {
                $.ajax({
                    url: '/api/cross_departmental_subjects',
                    type: 'POST',
                    contentType: 'application/json',
                    data: JSON.stringify({acad_groups: selectedGroups}),
                    success: function(crossSubjects) {
                        displayCrossDepartmentalSubjects(crossSubjects);
                    },
                    error: function() {
                        $('#crossDepartmentalContent').html('<p style="color: red;">Failed to load cross-departmental subjects.</p>');
                    }
                });
            }
            
            function displayCrossDepartmentalSubjects(crossSubjects) {
                const container = $('#crossDepartmentalContent');
                container.empty();
                
                if (!crossSubjects || crossSubjects.length === 0) {
                    container.html('<p style="color: #666; font-style: italic;">No cross-departmental subjects found in selected academic groups.</p>');
                    return;
                }
                
                container.append('<p><strong>Subjects available across multiple selected academic groups:</strong></p>');
                container.append('<p><em>Select cross-departmental subjects to load their courses:</em></p>');
                
                crossSubjects.forEach((subject, index) => {
                    container.append(`
                        <div class="cross-dept-subject">
                            <div class="cross-dept-checkbox">
                                <label>
                                    <input type="checkbox" name="crossDeptSubjects" value="${subject.Subject}" onchange="onCrossDeptSubjectChange()">
                                    <h5>${subject.Subject} - Cross-Departmental Subject</h5>
                                </label>
                            </div>
                            <div class="subject-info">
                                📊 ${subject.Course_Count} courses available in ${subject.Group_Count} academic groups<br>
                                📍 Available in: ${subject.Academic_Groups.join(', ')}
                            </div>
                        </div>
                    `);
                });
                
                // Add explanation
                container.append(`
                    <div style="margin-top: 15px; padding: 10px; background-color: #d1ecf1; border-radius: 5px;">
                        <small style="color: #0c5460;">
                            💡 <strong>Note:</strong> Cross-departmental subjects are available across multiple academic groups. 
                            Selecting them will load courses that can be taken by students from different departments.
                        </small>
                    </div>
                `);
            }
            
            function onCrossDeptSubjectChange() {
                // Get selected cross-departmental subjects
                const selectedCrossSubjects = $('input[name="crossDeptSubjects"]:checked').map(function() { 
                    return this.value; 
                }).get();
                
                // Get currently selected regular subject
                const regularSubject = $('#subjectSelect').val();
                
                // Combine and load courses
                const allSelectedSubjects = [];
                if (regularSubject) {
                    allSelectedSubjects.push(regularSubject);
                }
                allSelectedSubjects.push(...selectedCrossSubjects);
                
                if (allSelectedSubjects.length > 0) {
                    loadCoursesForMultipleSubjects(allSelectedSubjects);
                } else {
                    $('#courseSelection').html('<p>Please select a subject first...</p>');
                }
            }
            
            function loadCoursesForMultipleSubjects(subjects) {
                $('#courseSelection').html('<p>Loading courses...</p>');
                
                // Load courses for all selected subjects
                const coursePromises = subjects.map(subject => 
                    $.get(`/api/classes/${subject}`).catch(error => {
                        console.error(`Failed to load courses for ${subject}:`, error);
                        return { error: true, subject: subject, courses: [] };
                    })
                );
                
                Promise.all(coursePromises).then(results => {
                    const container = $('#courseSelection');
                    container.empty();
                    
                    let allCourses = [];
                    let subjectGroups = {};
                    let errorSubjects = [];
                    
                    // Group courses by subject
                    results.forEach((courses, index) => {
                        const subject = subjects[index];
                        if (courses.error) {
                            errorSubjects.push(subject);
                        } else if (courses && courses.length > 0) {
                            subjectGroups[subject] = courses;
                            allCourses.push(...courses);
                        }
                    });
                    
                    if (errorSubjects.length > 0) {
                        container.append(`
                            <div style="background-color: #f8d7da; color: #721c24; padding: 10px; margin-bottom: 15px; border-radius: 5px;">
                                <strong>⚠️ Warning:</strong> Failed to load courses for: ${errorSubjects.join(', ')}
                                <br><small>Please try again or select different subjects.</small>
                            </div>
                        `);
                    }
                    
                    if (allCourses.length === 0) {
                        container.append('<p>No courses found for selected subjects.</p>');
                        return;
                    }
                    
                    container.append('<p><strong>Select courses from all selected subjects:</strong></p>');
                    
                    // Display courses grouped by subject
                    Object.keys(subjectGroups).forEach(subject => {
                        const courses = subjectGroups[subject];
                        const isRegular = $('#subjectSelect').val() === subject;
                        const subjectType = isRegular ? 'Regular Subject' : 'Cross-Departmental Subject';
                        const subjectColor = isRegular ? '#007bff' : '#28a745';
                        
                        container.append(`
                            <h4 style="color: ${subjectColor}; margin-top: 15px; margin-bottom: 8px;">
                                ${subject} - ${subjectType} (${courses.length} courses)
                            </h4>
                        `);
                        
                        courses.forEach(course => {
                            container.append(`
                                <div class="checkbox-item">
                                    <label>
                                        <input type="checkbox" name="courses" value="${course.Course_Code}">
                                        ${course.Course_Code} - ${course.Course_Title}
                                    </label>
                                </div>
                            `);
                        });
                    });
                    
                    // Add summary
                    if (Object.keys(subjectGroups).length > 1) {
                        container.append(`
                            <div style="margin-top: 15px; padding: 10px; background-color: #d1ecf1; border-radius: 5px;">
                                <small style="color: #0c5460;">
                                    📊 <strong>Summary:</strong> Loaded ${allCourses.length} courses from ${Object.keys(subjectGroups).length} subjects.
                                    You can select courses from both regular and cross-departmental subjects.
                                </small>
                            </div>
                        `);
                    }
                    
                }).catch(error => {
                    console.error('Failed to load courses for multiple subjects:', error);
                    $('#courseSelection').html(`
                        <p style="color: red;">
                            <strong>Error:</strong> Failed to load courses. 
                            <br>Please check your connection and try again.
                            <br><small>Error details: ${error.message || error}</small>
                        </p>
                    `);
                });
            }
            
            function clearAcadGroupSelection() {
                $('input[name="acadGroups"]').prop('checked', false);
                $('input[name="crossDeptSubjects"]').prop('checked', false);
                $('#subjectSelect').empty().append('<option value="">Choose Subject...</option>').prop('disabled', true);
                $('#subjectSelectHint').text('Please select a single academic group to enable subject selection').show();
                $('#crossDepartmentalSection').hide();
                $('#courseSelection').html('<p>Please select a subject first...</p>');
                $('#teacherSelection').html('<p>No courses selected...</p>');
                $('#roomSelection').html('<p>Loading rooms...</p>');
                loadRooms(); // Reload default rooms
                updateSelectionCounts(); // Update count
            }
            
            function loadSubjects() {
                // This function is now replaced by loadSubjectsFromSelectedGroups
                // Keep for backward compatibility but show message
                console.log('loadSubjects() is deprecated. Use loadSubjectsFromSelectedGroups() instead.');
                
                // Get selected groups from checkboxes
                const selectedGroups = $('input[name="acadGroups"]:checked').map(function() { 
                    return this.value; 
                }).get();
                
                if (selectedGroups.length > 0) {
                    loadSubjectsFromSelectedGroups();
                } else {
                    $('#subjectSelect').empty().append('<option value="">Please select academic groups first...</option>');
                }
            }
            
            function loadCourses() {
                const subject = $('#subjectSelect').val();
                if (!subject) {
                    $('#courseSelection').html('<p>Please select a subject first...</p>');
                    return;
                }
                
                // Get selected cross-departmental subjects
                const selectedCrossSubjects = $('input[name="crossDeptSubjects"]:checked').map(function() { 
                    return this.value; 
                }).get();
                
                // Combine regular and cross-departmental subjects
                const allSelectedSubjects = [subject, ...selectedCrossSubjects];
                
                loadCoursesForMultipleSubjects(allSelectedSubjects);
            }
            
            function updateClassSections() {
                const selectedCourses = $('input[name="courses"]:checked').map(function() { 
                    return this.value; 
                }).get();
                
                if (selectedCourses.length === 0) {
                    $('#teacherSelection').html('<p>No courses selected...</p>');
                    $('#roomSelection').html('<p>No courses selected...</p>');
                    updateSelectionCounts(); // Update count
                    return;
                }
                
                // Load all class sections for selected courses
                $.ajax({
                    url: '/api/classes_by_codes',
                    type: 'POST',
                    contentType: 'application/json',
                    data: JSON.stringify({course_codes: selectedCourses}),
                    success: function(classData) {
                        // Store class data globally for later use
                        window.currentClassSections = classData;
                        
                        // Load teachers for these classes
                        const classNbrs = classData.map(cls => cls.Class_Nbr);
                    if (classNbrs.length > 0) {
                        loadTeachersForClasses(classNbrs);
                    }
                        
                        // Display class sections info (optional)
                        console.log(`Loaded ${classData.length} class sections for ${selectedCourses.length} courses`);
                    },
                    error: function() {
                        console.error('Failed to load class sections');
                    }
                });
                
                // Load rooms for selected courses
                $.ajax({
                    url: '/api/smart_rooms_for_courses',
                    type: 'POST',
                    contentType: 'application/json',
                    data: JSON.stringify({course_codes: selectedCourses}),
                    success: function(data) {
                        updateRoomSelection(data, selectedCourses);
                    },
                    error: function() {
                        console.error('Failed to load rooms');
                        // Fallback to loading all rooms
                        loadRooms();
                    }
                });
            }
            
            function updateRoomSelection(roomData, selectedCourses) {
                const container = $('#roomSelection');
                container.empty();
                
                const rooms = roomData.rooms || [];
                const matchInfo = roomData.match_info || {};
                
                if (rooms.length === 0) {
                    container.html('<p>No suitable rooms found.</p>');
                    return;
                }
                
                // Display match information in English
                let matchInfoHtml = '<div style="background-color: #e3f2fd; padding: 10px; border-radius: 5px; margin-bottom: 10px;">';
                matchInfoHtml += `<strong> Room Matching Results (${selectedCourses.length} courses):</strong><br>`;
                
                const requirements = [];
                if (matchInfo.needs_lab) requirements.push('🧪 Laboratory Required');
                if (matchInfo.needs_computer_lab) requirements.push('💻 Computer Lab Required');
                if (matchInfo.needs_engineering_lab) requirements.push('⚙️ Engineering Lab Required');
                if (matchInfo.needs_special_facility) requirements.push('🏛️ Special Facility Required');
                
                if (requirements.length > 0) {
                    matchInfoHtml += `Requirements Analysis: ${requirements.join(', ')}<br>`;
                } else {
                    matchInfoHtml += 'Requirements Analysis: Regular classroom suitable<br>';
                }
                
                matchInfoHtml += `Found ${rooms.length} suitable rooms`;
                matchInfoHtml += '</div>';
                
                container.append(matchInfoHtml);
                
                // Group rooms by type
                const labRooms = rooms.filter(room => 
                    room.Description && room.Description.toLowerCase().includes('lab'));
                const classrooms = rooms.filter(room => 
                    room.Description && (room.Description.toLowerCase().includes('classroom') || 
                                        room.Description.toLowerCase().includes('class')));
                const otherRooms = rooms.filter(room => 
                    !room.Description || (!room.Description.toLowerCase().includes('lab') && 
                                          !room.Description.toLowerCase().includes('classroom') &&
                                          !room.Description.toLowerCase().includes('class')));
                
                // Display labs first
                if (labRooms.length > 0) {
                    container.append('<h4 style="color: #1976d2; margin-top: 15px;">🧪 Laboratories (' + labRooms.length + ' rooms)</h4>');
                    labRooms.forEach(room => {
                        container.append(createRoomCheckbox(room, 'lab'));
                    });
                }
                
                // Then display classrooms
                if (classrooms.length > 0) {
                    container.append('<h4 style="color: #388e3c; margin-top: 15px;">🏫 Classrooms (' + classrooms.length + ' rooms)</h4>');
                    classrooms.slice(0, 15).forEach(room => {  // Limit to 15 classrooms
                        container.append(createRoomCheckbox(room, 'classroom'));
                    });
                    if (classrooms.length > 15) {
                        container.append(`<p style="color: #666; font-style: italic;">... ${classrooms.length - 15} more classrooms not displayed</p>`);
                    }
                }
                
                // Finally other rooms
                if (otherRooms.length > 0) {
                    container.append('<h4 style="color: #f57c00; margin-top: 15px;">🏢 Other Facilities (' + otherRooms.length + ' rooms)</h4>');
                    otherRooms.slice(0, 5).forEach(room => {  // Limit to 5 other rooms
                        container.append(createRoomCheckbox(room, 'other'));
                    });
                }
                
                updateSelectionCounts(); // Update room selection count
            }
            
            function createRoomCheckbox(room, type) {
                const typeIcon = {
                    'lab': '🧪',
                    'classroom': '🏫', 
                    'other': '🏢'
                }[type] || '📍';
                
                return `
                    <div class="checkbox-item">
                        <label>
                            <input type="checkbox" name="rooms" value="${room.Room_ID}">
                            ${typeIcon} ${room.Room_ID}: ${room.Description || 'No description'} 
                            <span style="color: #666;">(Capacity: ${room.Capacity || 'N/A'})</span>
                        </label>
                    </div>
                `;
            }
            
            function loadTeachersForClasses(classNbrs) {
                $.ajax({
                    url: '/api/teachers_for_classes',
                    type: 'POST',
                    contentType: 'application/json',
                    data: JSON.stringify({class_nbrs: classNbrs}),
                    success: function(data) {
                        const container = $('#teacherSelection');
                        container.empty();
                        
                        if (data.length === 0) {
                            container.html('<p>No teachers found for selected courses.</p>');
                            updateSelectionCounts(); // Update count
                            return;
                        }
                        
                        const uniqueTeachers = {};
                        data.forEach(teacher => {
                            uniqueTeachers[teacher.F_ID] = teacher;
                        });
                        
                        Object.values(uniqueTeachers).forEach(teacher => {
                            container.append(`
                                <div class="checkbox-item">
                                    <label>
                                        <input type="checkbox" name="teachers" value="${teacher.F_ID}">
                                        ${teacher.First_Name} ${teacher.Last_Name} (${teacher.F_ID})
                                    </label>
                                </div>
                            `);
                        });
                        updateSelectionCounts(); // Update count
                    },
                    error: function() {
                        $('#teacherSelection').html('<p style="color: red;">Failed to load teachers.</p>');
                        updateSelectionCounts(); // Update count
                        console.error('Failed to load teachers');
                    }
                });
            }
            
            function loadRooms() {
                $.get('/api/rooms', function(data) {
                    const container = $('#roomSelection');
                    container.empty();
                    
                    if (data.length === 0) {
                        container.html('<p>No rooms available.</p>');
                        updateSelectionCounts(); // Update count
                        return;
                    }
                    
                    data.slice(0, 20).forEach(room => {  // Show first 20 rooms
                        container.append(`
                            <div class="checkbox-item">
                                <label>
                                    <input type="checkbox" name="rooms" value="${room.Room_ID}">
                                    ${room.Room_ID}: ${room.Description || 'No description'} (Cap: ${room.Capacity})
                                </label>
                            </div>
                        `);
                    });
                    updateSelectionCounts(); // Update count
                }).fail(function() {
                    $('#roomSelection').html('<p style="color: red;">Failed to load rooms.</p>');
                    updateSelectionCounts(); // Update count
                    console.error('Failed to load rooms');
                });
            }
            
            // Time slot management functionality
            function loadTimeSlotStatus() {
                $.get('/api/get_time_slot_status', function(data) {
                    timeSlotStatus = data;
                    currentDisabledSlots = new Set(data.disabled_slots);
                    generateTimeSlotTable();
                }).fail(function() {
                    console.error('Failed to load time slot status');
                });
            }
            
            function generateTimeSlotTable() {
                const days = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
                const periods = [
                    {start: "08:00", end: "09:15", name: "Period 1"},
                    {start: "09:30", end: "10:45", name: "Period 2"},
                    {start: "11:00", end: "12:15", name: "Period 3"},
                    {start: "13:00", end: "14:15", name: "Period 4"},
                    {start: "14:30", end: "15:45", name: "Period 5"},
                    {start: "16:00", end: "17:15", name: "Period 6"},
                    {start: "17:30", end: "18:45", name: "Period 7"}
                ];
                
                const tbody = $('#timeslotTable tbody');
                tbody.empty();
                
                periods.forEach(period => {
                    const row = $('<tr></tr>');
                    row.append(`<td><strong>${period.start}-${period.end}</strong><br><small>${period.name}</small></td>`);
                    
                    days.forEach(day => {
                        const timeId = `${day}_${period.start}-${period.end}`;
                        const isDisabled = currentDisabledSlots.has(timeId);
                        const cellClass = isDisabled ? 'timeslot-disabled' : 'timeslot-available';
                        const cellText = isDisabled ? 'Disabled' : 'Available';
                        
                        const cell = $(`<td class="timeslot-cell ${cellClass}" data-time-id="${timeId}">${cellText}</td>`);
                        cell.click(function() {
                            toggleTimeSlot(timeId);
                        });
                        
                        row.append(cell);
                    });
                    
                    tbody.append(row);
                });
            }
            
            function toggleTimeSlot(timeId) {
                if (currentDisabledSlots.has(timeId)) {
                    currentDisabledSlots.delete(timeId);
                } else {
                    currentDisabledSlots.add(timeId);
                }
                
                const cell = $(`[data-time-id="${timeId}"]`);
                if (currentDisabledSlots.has(timeId)) {
                    cell.removeClass('timeslot-available').addClass('timeslot-disabled').text('Disabled');
                } else {
                    cell.removeClass('timeslot-disabled').addClass('timeslot-available').text('Available');
                }
            }
            
            function saveTimeSlotSettings() {
                const disabledSlots = Array.from(currentDisabledSlots);
                
                $.ajax({
                    url: '/api/disable_time_slots',
                    type: 'POST',
                    contentType: 'application/json',
                    data: JSON.stringify({time_slot_patterns: disabledSlots}),
                    success: function(data) {
                        if (data.success) {
                            $('#results').html(`
                                <h3>Time Slot Settings Saved</h3>
                                <p><strong>Disabled Time Slots:</strong> ${data.disabled_count}</p>
                            `).show();
                        } else {
                            $('#results').html(`<p style="color: red;">Error saving settings: ${data.error}</p>`).show();
                        }
                    },
                    error: function() {
                        $('#results').html('<p style="color: red;">Failed to save time slot settings.</p>').show();
                    }
                });
            }
            
            function enableAllTimeSlots() {
                $.ajax({
                    url: '/api/enable_all_time_slots',
                    type: 'POST',
                    success: function(data) {
                        if (data.success) {
                            currentDisabledSlots.clear();
                            generateTimeSlotTable();
                            $('#results').html(`
                                <h3>Time Slot Settings</h3>
                                <p><strong>All Time Slots Enabled:</strong> ${data.message}</p>
                            `).show();
                        } else {
                            $('#results').html(`<p style="color: red;">Error: ${data.error}</p>`).show();
                        }
                    },
                    error: function() {
                        $('#results').html('<p style="color: red;">Failed to enable all time slots.</p>').show();
                    }
                });
            }
            
            function generateSchedule() {
                // Get selected course codes
                const selectedCourses = $('input[name="courses"]:checked').map(function() { 
                    return this.value; 
                }).get();
                
                if (selectedCourses.length === 0) {
                    alert('Please select at least one course.');
                    return;
                }
                
                // Get class numbers from stored class sections data
                const classNumbers = window.currentClassSections ? 
                    window.currentClassSections.map(cls => cls.Class_Nbr) : [];
                
                if (classNumbers.length === 0) {
                    alert('No class sections found for selected courses. Please reselect courses.');
                    return;
                }
                
                // Get selected academic groups (multi-select)
                const selectedAcadGroups = $('input[name="acadGroups"]:checked').map(function() { 
                    return this.value; 
                }).get();
                
                const selections = {
                    term: $('#termSelect').val(),
                    session: $('#sessionSelect').val(),
                    campus: $('#campusSelect').val(),
                    acad_group: selectedAcadGroups.length > 0 ? selectedAcadGroups[0] : '', // Use first group for compatibility
                    acad_groups: selectedAcadGroups, // Send all selected groups
                    subject: $('#subjectSelect').val(),
                    classes: classNumbers,  // Use class numbers instead of course codes
                    teachers: $('input[name="teachers"]:checked').map(function() { return this.value; }).get(),
                    rooms: $('input[name="rooms"]:checked').map(function() { return this.value; }).get()
                };
                
                $('#results').html('<p>Generating schedule...</p>').show();
                
                $.ajax({
                    url: '/api/generate_schedule',
                    type: 'POST',
                    contentType: 'application/json',
                    data: JSON.stringify(selections),
                    success: function(data) {
                        if (data.success) {
                            $('#results').html(`
                                <h3>Schedule Generation Results</h3>
                                <p><strong>Selected Academic Groups:</strong> ${selectedAcadGroups.join(', ')}</p>
                                <p><strong>Selected Courses:</strong> ${selectedCourses.length} courses</p>
                                <p><strong>Class Sections:</strong> ${classNumbers.length} sections</p>
                                <p><strong>Successfully Scheduled:</strong> ${data.scheduled_count} classes</p>
                                <p><strong>Conflicts:</strong> ${data.conflict_count} classes</p>
                                <p><strong>Available Time Slots:</strong> ${data.available_time_slots}</p>
                            `);
                            
                            // Enable export schedule results button
                            $('#exportScheduleBtn').prop('disabled', false).css('background-color', '#28a745');
                            
                            displayTimetable(data.timetable);
                        } else {
                            $('#results').html(`<p style="color: red;">Error: ${data.error}</p>`);
                        }
                    },
                    error: function() {
                        $('#results').html('<p style="color: red;">Failed to generate schedule.</p>');
                    }
                });
            }
            
            function displayTimetable(timetable) {
                const days = ["Sunday","Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
                const times = ["08:00:00-09:15:00", "09:30:00-10:45:00", "11:00:00-12:15:00", "13:00:00-14:15:00", "14:30:00-15:45:00", "16:00:00-17:15:00", "17:30:00-18:45:00"];
                
                let html = '<h3>Generated Timetable</h3><table class="timetable"><tr><th>Time</th>';
                days.forEach(day => {
                    html += `<th>${day}</th>`;
                });
                html += '</tr>';
                
                times.forEach(time => {
                    html += `<tr><td><strong>${time}</strong></td>`;
                    days.forEach(day => {
                        html += '<td>';
                        if (timetable[day] && timetable[day][time]) {
                            timetable[day][time].forEach(cls => {
                                html += `<div class="class-item">
                                    <strong>${cls.course_code}</strong><br>
                                    ${cls.section}<br>
                                    ${cls.teacher}<br>
                                    Room: ${cls.room}
                                </div>`;
                            });
                        }
                        html += '</td>';
                    });
                    html += '</tr>';
                });
                
                html += '</table>';
                $('#timetable').html(html);
            }
            
            function simulateImport() {
                const fileInput = document.getElementById('excelFileInput');
                const file = fileInput.files[0];
                
                if (!file) {
                    alert('Please select an Excel file first.');
                    return;
                }
                
                const formData = new FormData();
                formData.append('file', file);
                
                $('#results').html('<p>Simulating import...</p>').show();
                
                $.ajax({
                    url: '/api/import_excel',
                    type: 'POST',
                    data: formData,
                    processData: false,
                    contentType: false,
                    success: function(data) {
                        if (data.success) {
                            $('#results').html(`
                                <h3>Simulated Import Results</h3>
                                <p><strong>Message:</strong> ${data.message}</p>
                                <p><strong>File Info:</strong> ${data.file_info}</p>
                                <h4>Imported Data Summary:</h4>
                                <ul>
                                    <li><strong>Faculty:</strong> ${data.imported_data.Faculty}</li>
                                    <li><strong>Rooms:</strong> ${data.imported_data.Room}</li>
                                    <li><strong>Courses:</strong> ${data.imported_data.Course}</li>
                                    <li><strong>Campus:</strong> ${data.imported_data.Campus}</li>
                                    <li><strong>Students:</strong> ${data.imported_data.Students}</li>
                                    <li><strong>Class Sections:</strong> ${data.imported_data.Class_Sections}</li>
                                </ul>
                            `);
                        } else {
                            $('#results').html(`<p style="color: red;">Import failed: ${data.error}</p>`);
                        }
                    },
                    error: function() {
                        $('#results').html('<p style="color: red;">Failed to simulate import.</p>');
                    }
                });
            }
            
            function exportScheduleResults() {
                // Check if button is available
                if ($('#exportScheduleBtn').prop('disabled')) {
                    alert('Please generate a timetable first before exporting schedule results.');
                    return;
                }
                
                $('#results').html('<p>Exporting schedule results...</p>').show();
                
                $.get('/api/export_schedule_results', function(data) {
                    if (data.success) {
                        lastGeneratedFileName = data.filename;  // 保存文件名
                        $('#results').html(`
                            <h3>Schedule Results Exported Successfully</h3>
                            <p><strong>     File:</strong> ${data.filename}</p>
                            <p><strong>     Format:</strong> Excel (.xlsx) with 36 standard columns</p>
                            <p><strong>     Total Records:</strong> ${data.record_count}</p>
                            <p><strong>     Scheduled Sessions:</strong> ${data.scheduled_count}</p>
                            <p><strong>     Conflicts:</strong> ${data.conflicts_count}</p>
                            <p style="color: #28a745; font-weight: bold;">Ready for download!</p>
                        `);
                        
                        $('#fileName').text(data.filename);
                        $('#lastGeneratedFile').show();
                    } else {
                        $('#results').html(`<p style="color: red;">Export failed: ${data.error}</p>`);
                    }
                }).fail(function() {
                    $('#results').html('<p style="color: red;">Failed to export schedule results.</p>');
                });
            }
            
            function downloadFile() {
                if (!lastGeneratedFileName) {
                    alert('No file available for download. Please export data first.');
                    return;
                }
                
                // Use simpler direct download method
                const downloadUrl = `/download/${encodeURIComponent(lastGeneratedFileName)}`;
                
                // Open download link directly
                window.open(downloadUrl, '_blank');
                
                $('#results').html(`
                    <h3>File Download Started</h3>
                    <p><strong>File:</strong> ${lastGeneratedFileName}</p>
                    <p>Download should start automatically. If the file doesn't open properly, please check if you have Excel installed.</p>
                `).show();
            }
        </script>
    </body>
    </html>
    """
    
    @app.route('/')
    def index():
        return render_template_string(HTML_TEMPLATE)
    
    @app.route('/api/terms')
    def get_terms():
        return jsonify(scheduler.get_terms())
    
    @app.route('/api/sessions')
    def get_sessions():
        return jsonify(scheduler.get_sessions())
    
    @app.route('/api/campuses')
    def get_campuses():
        return jsonify(scheduler.get_campuses())
    
    @app.route('/api/acad_groups')
    def get_acad_groups():
        return jsonify(scheduler.get_acad_groups())
    
    @app.route('/api/subjects/<acad_group>')
    def get_subjects(acad_group):
        return jsonify(scheduler.get_subjects_by_acad_group(acad_group))
    
    @app.route('/api/classes/<subject>')
    def get_classes(subject):
        return jsonify(scheduler.get_classes_by_subject(subject))
    
    @app.route('/api/teachers_for_classes', methods=['POST'])
    def get_teachers_for_classes():
        class_nbrs = request.json.get('class_nbrs', [])
        return jsonify(scheduler.get_teachers_for_classes(class_nbrs))
    
    @app.route('/api/rooms')
    def get_rooms():
        return jsonify(scheduler.get_available_rooms())
    
    @app.route('/api/generate_schedule', methods=['POST'])
    def generate_schedule():
        selections = request.json
        scheduler.set_selections(**selections)
        result = scheduler.generate_timetable()
        return jsonify(result)
    
    @app.route('/api/import_excel', methods=['POST'])
    def import_excel():
        """Import Excel file endpoint"""
        if 'file' not in request.files:
            return jsonify({'success': False, 'error': 'No file provided'})
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({'success': False, 'error': 'No file selected'})
        
        if file and file.filename.endswith(('.xlsx', '.xls')):
            try:
                # Save uploaded file temporarily
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                temp_filename = f"temp_import_{timestamp}.xlsx"
                file.save(temp_filename)
                
                # Process the file
                result = scheduler.import_excel_data(temp_filename)
                
                # Clean up temp file
                if os.path.exists(temp_filename):
                    os.remove(temp_filename)
                
                return jsonify(result)
            except Exception as e:
                return jsonify({'success': False, 'error': str(e)})
        else:
            return jsonify({'success': False, 'error': 'Invalid file format. Please upload .xlsx or .xls file'})
    
    @app.route('/api/export_schedule_results')
    def export_schedule_results():
        """Export latest schedule results to Excel endpoint"""
        try:
            result = scheduler.export_schedule_results_to_excel()
            if result['success']:
                return jsonify({
                    'success': True, 
                    'filename': result['filename'],
                    'record_count': result['record_count'],
                    'scheduled_count': result.get('scheduled_count', 0),
                    'conflicts_count': result.get('conflicts_count', 0),
                    'download_url': f'/download/{result["filename"]}'
                })
            else:
                return jsonify(result)
        except Exception as e:
            return jsonify({'success': False, 'error': str(e)})
    
    @app.route('/api/schedule_results_status')
    def get_schedule_results_status():
        """Get current schedule results status"""
        return jsonify(scheduler.get_schedule_results_status())
    
    @app.route('/download/<filename>')
    def download_file(filename):
        """Download generated Excel files"""
        from flask import send_file, abort, current_app
        import os
        import traceback
        
        try:
            print(f"Download request for file: {filename}")
            
            # Security check: only allow downloading files in current directory
            if not filename or '..' in filename or '/' in filename or '\\' in filename:
                print(f"Security check failed for filename: {filename}")
                abort(400, description="Invalid filename")
            
            # Get full file path
            file_path = os.path.abspath(filename)
            print(f"Full file path: {file_path}")
            
            # Check if file exists
            if not os.path.exists(file_path):
                print(f"File not found: {file_path}")
                # List current directory files for debugging
                current_files = os.listdir('.')
                excel_files = [f for f in current_files if f.endswith(('.xlsx', '.xls'))]
                print(f"Available Excel files in current directory: {excel_files}")
                abort(404, description="File not found")
            
            # Check if file is Excel file
            if not filename.endswith(('.xlsx', '.xls')):
                print(f"Invalid file type: {filename}")
                abort(400, description="Invalid file type")
            
            # Check file size is reasonable
            file_size = os.path.getsize(file_path)
            if file_size == 0:
                print(f"File is empty: {filename}")
                abort(500, description="File is empty or corrupted")
            
            print(f"Sending file: {filename}, size: {file_size} bytes")
            
            # Send file with more explicit parameters
            return send_file(
                file_path,
                as_attachment=True,
                download_name=filename,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
        except Exception as e:
            print(f"Error downloading file: {str(e)}")
            print("Full traceback:")
            traceback.print_exc()
            abort(500, description=f"Download failed: {str(e)}")
    
    # Add another API route to support different URL formats
    @app.route('/api/download/<filename>')
    def api_download_file(filename):
        """API version of download endpoint"""
        return download_file(filename)
    
    @app.route('/api/disable_time_slots', methods=['POST'])
    def disable_time_slots():
        """Disable specified time slots"""
        patterns = request.json.get('time_slot_patterns', [])
        result = scheduler.disable_time_slots(patterns)
        return jsonify(result)
    
    @app.route('/api/enable_all_time_slots', methods=['POST'])
    def enable_all_time_slots():
        """Enable all time slots"""
        result = scheduler.enable_all_time_slots()
        return jsonify(result)
    
    @app.route('/api/get_time_slot_status')
    def get_time_slot_status():
        """Get time slot status"""
        all_slots = scheduler._generate_time_slots()
        available_slots = scheduler.get_available_time_slots()
        disabled_slots = list(scheduler.disabled_time_slots)
        
        return jsonify({
            'total_slots': len(all_slots),
            'available_slots': len(available_slots),
            'disabled_slots': disabled_slots,
            'disabled_count': len(disabled_slots)
        })
    
    @app.route('/api/classes_by_codes', methods=['POST'])
    def get_classes_by_codes():
        course_codes = request.json.get('course_codes', [])
        return jsonify(scheduler.get_classes_by_course_codes(course_codes))
    
    @app.route('/api/smart_rooms_for_courses', methods=['POST'])
    def get_smart_rooms_for_courses():
        course_codes = request.json.get('course_codes', [])
        result = scheduler.get_smart_rooms_for_courses(course_codes)
        return jsonify(result)
    
    @app.route('/api/subjects_multi_groups', methods=['POST'])
    def get_subjects_multi_groups():
        acad_groups = request.json.get('acad_groups', [])
        return jsonify(scheduler.get_subjects_by_multiple_acad_groups(acad_groups))
    
    @app.route('/api/cross_departmental_subjects', methods=['POST'])
    def get_cross_departmental_subjects():
        acad_groups = request.json.get('acad_groups', [])
        return jsonify(scheduler.get_cross_departmental_subjects(acad_groups))
    
    return app

if __name__ == "__main__":
    # Create and run the web application
    app = create_web_api()
    print("Starting Web Scheduling System...")
    print("Open your browser and go to: http://localhost:5100")
    # Check if running in production mode
    debug_mode = os.environ.get('FLASK_DEBUG', '0') == '1'
    app.run(debug=debug_mode, host='0.0.0.0', port=5100) 
