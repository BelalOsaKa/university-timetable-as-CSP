"""
Data loading and preprocessing module for timetable CSP
"""
import pandas as pd
import re
from typing import Dict, List, Set, Tuple, Any

class DataLoader:
    def __init__(self):
        self.courses = {}
        self.instructors = {}
        self.rooms = {}
        self.sections = {}
        self.timeslots = {}
        
    def load_courses(self, filepath: str) -> Dict[str, Dict]:
        """Load courses from CSV file"""
        df = pd.read_csv(filepath)
        courses = {}
        
        for _, row in df.iterrows():
            course_id = row['CourseID']
            course_type = 'lab' if 'Lab' in row['Type'] else 'lec'
            courses[course_id] = {
                'name': row['CourseName'],
                'credits': row['Credits'],
                'type': course_type
            }
        self.courses = courses
        return courses
    
    def load_instructors(self, filepath: str) -> Dict[str, Dict]:
        """Load instructors from CSV file"""
        df = pd.read_csv(filepath)
        instructors = {}
        
        for _, row in df.iterrows():
            instructor_id = row['InstructorID']
            qualified_courses = set()
            
            # Parse qualified courses
            if pd.notna(row['QualifiedCourses']):
                courses_str = str(row['QualifiedCourses'])
                qualified_courses = set(courses_str.split(','))
                qualified_courses = {c.strip() for c in qualified_courses if c.strip()}
            
            # Parse preferred slots
            preferred_slots = set()
            if pd.notna(row['PreferredSlots']):
                pref_str = str(row['PreferredSlots'])
                if pref_str != 'Any time':
                    # Extract day restrictions
                    day_mapping = {
                        'Sunday': 0, 'Monday': 1, 'Tuesday': 2, 'Wednesday': 3, 
                        'Thursday': 4, 'Friday': 5, 'Saturday': 6
                    }
                    for day, day_num in day_mapping.items():
                        if day in pref_str:
                            # Add all timeslots for this day as restricted
                            for ts_id in range(day_num * 4, (day_num + 1) * 4):
                                preferred_slots.add(ts_id)
            
            instructors[instructor_id] = {
                'name': row['Name'],
                'role': row['Role'],
                'qualified_courses': qualified_courses,
                'preferred_slots': preferred_slots,  # slots to avoid
                'available_slots': set(range(20)) - preferred_slots  # available slots
            }
        self.instructors = instructors
        return instructors
    
    def load_rooms(self, filepath: str) -> Dict[str, Dict]:
        """Load rooms from CSV file"""
        df = pd.read_csv(filepath)
        rooms = {}
        
        for _, row in df.iterrows():
            room_id = row['RoomID']
            room_type = 'lab' if row['Type'] == 'Lab' else 'lec'
            rooms[room_id] = {
                'type': room_type,
                'capacity': row['Capacity']
            }
        self.rooms = rooms
        return rooms
    
    def load_timeslots(self, filepath: str) -> Dict[int, Dict]:
        """Load timeslots from CSV file"""
        df = pd.read_csv(filepath)
        timeslots = {}
        
        day_mapping = {
            'Sunday': 0, 'Monday': 1, 'Tuesday': 2, 'Wednesday': 3, 
            'Thursday': 4, 'Friday': 5, 'Saturday': 6
        }
        
        for idx, row in df.iterrows():
            timeslot_id = idx
            day_num = day_mapping[row['Day']]
            timeslots[timeslot_id] = {
                'day': row['Day'],
                'day_num': day_num,
                'start_time': row['StartTime'],
                'end_time': row['EndTime'],
                'timeslot_id': row['TimeSlotID']
            }
        self.timeslots = timeslots
        return timeslots
    
    def load_sections(self, filepath: str) -> Dict[str, Dict]:
        """Load sections from CSV file"""
        df = pd.read_csv(filepath)
        sections = {}
        
        for _, row in df.iterrows():
            section_id = row['SectionID']
            student_count = row['StudentCount']
            
            # Parse courses for this section
            courses_str = str(row['Courses'])
            course_list = [c.strip() for c in courses_str.split(',') if c.strip()]
            
            sections[section_id] = {
                'student_count': student_count,
                'courses': course_list
            }
        self.sections = sections
        return sections
    
    def load_all_data(self, base_path: str = ".") -> Dict[str, Any]:
        """Load all data files"""
        try:
            courses = self.load_courses(f"{base_path}/Courses.csv")
            instructors = self.load_instructors(f"{base_path}/Instructor.csv")
            rooms = self.load_rooms(f"{base_path}/Rooms.csv")
            timeslots = self.load_timeslots(f"{base_path}/TimeSlots.csv")
            sections = self.load_sections(f"{base_path}/Sections.csv")
            
            return {
                'courses': courses,
                'instructors': instructors,
                'rooms': rooms,
                'timeslots': timeslots,
                'sections': sections
            }
        except Exception as e:
            print(f"Error loading data: {e}")
            return {}
    
    def get_statistics(self) -> Dict[str, Any]:
        """Get statistics about loaded data"""
        stats = {
            'total_courses': len(self.courses),
            'total_instructors': len(self.instructors),
            'total_rooms': len(self.rooms),
            'total_timeslots': len(self.timeslots),
            'total_sections': len(self.sections),
            'lecture_rooms': len([r for r in self.rooms.values() if r['type'] == 'lec']),
            'lab_rooms': len([r for r in self.rooms.values() if r['type'] == 'lab']),
            'lecture_courses': len([c for c in self.courses.values() if c['type'] == 'lec']),
            'lab_courses': len([c for c in self.courses.values() if c['type'] == 'lab'])
        }
        return stats

if __name__ == "__main__":
    loader = DataLoader()
    data = loader.load_all_data()
    
    if data:
        print("Data loaded successfully!")
        stats = loader.get_statistics()
        for key, value in stats.items():
            print(f"{key}: {value}")
    else:
        print("Failed to load data")
