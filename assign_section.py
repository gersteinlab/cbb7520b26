import pandas as pd
import numpy as np
from collections import defaultdict

# Read the Excel file
df = pd.read_excel('student_info.xlsx')

# Define the 5 section times
section_times = [
    "Thu 11:35-12:25 pm",
    "Thu 3:30-4:20 pm",
    "Thu 7:00-7:50 pm",
    "Fri 10:30-11:20 am",
    "Fri 1:30-2:20 pm"
]

# Extract relevant columns
availability_col = df.columns[13]
name_col = df.columns[1]
email_col = df.columns[2]
enrollment_col = df.columns[4]
undergrad_major_col = df.columns[5]  # Column F - undergrad major
class_year_col = df.columns[6]
grad_program_col = df.columns[7]  # Column H - grad program

# Create a clean dataframe for processing
# For major/program: use column F for undergrads, column H for grad students
def get_major_or_program(row):
    enrollment = str(row[enrollment_col]).upper()
    if 'UNDERGRADUATE' in enrollment:
        return row[undergrad_major_col]
    else:  # Graduate students (PhD, Master, etc.)
        return row[grad_program_col]

students = pd.DataFrame({
    'Name': df[name_col],
    'Email': df[email_col],
    'Enrollment': df[enrollment_col],
    'Major': df.apply(get_major_or_program, axis=1),
    'Class_Year': df[class_year_col],
    'Availability': df[availability_col]
})

# Parse availability for each student
def parse_availability(avail_str):
    """Extract which section times a student is available for"""
    if pd.isna(avail_str):
        return []
    available_times = []
    for time in section_times:
        if time in str(avail_str):
            available_times.append(time)
    return available_times

students['Available_Times'] = students['Availability'].apply(parse_availability)

# Separate students with and without availability
students_with_avail = students[students['Available_Times'].apply(len) > 0].reset_index(drop=True)
students_without_avail = students[students['Available_Times'].apply(len) == 0].reset_index(drop=True)

print(f"Total students in spreadsheet: {len(students)}")
print(f"Students with availability data: {len(students_with_avail)}")
print(f"Students without availability data: {len(students_without_avail)}")
if len(students_without_avail) > 0:
    print("\nStudents without availability:")
    for idx, student in students_without_avail.iterrows():
        print(f"  - {student['Name']} ({student['Email']})")

print(f"\nEnrollment breakdown (students with availability):")
print(students_with_avail['Enrollment'].value_counts())
print(f"\nMajor/Program breakdown (students with availability):")
print(students_with_avail['Major'].value_counts())

# Use students_with_avail for the rest of the processing
students = students_with_avail

# Categorize majors/programs into broader groups for better balancing
def categorize_major(major):
    if pd.isna(major):
        return "Unknown"
    major_str = str(major).upper()
    
    # Computer Science / Computational
    if 'CS' in major_str or 'COMPUTER' in major_str or 'COMP' in major_str:
        return "CS/Computational"
    
    # Statistics and Data Science
    elif 'S&DS' in major_str or 'STATS' in major_str or 'DATA' in major_str:
        return "Stats/Data Science"
    
    # Computational Biology / Bioinformatics (CBB, BQBS, BIS)
    elif 'CBB' in major_str or 'BQBS' in major_str or 'BIS' in major_str or 'BIOINF' in major_str or 'HEALTH INF' in major_str:
        return "Comp Bio/Bioinformatics"
    
    # Biomedical Engineering
    elif 'BENG' in major_str or 'BME' in major_str:
        return "Biomedical Engineering"
    
    # Life Sciences (MB&B, MCDB, Neuroscience, PMB, etc.)
    elif any(term in major_str for term in ['MB&B', 'MCDB', 'NEURO', 'PMB', 'MCGD', 'PEB', 'BBS']):
        return "Life Sciences"
    
    # Public Health / Epidemiology
    elif 'MPH' in major_str or 'EMD' in major_str or 'YSPH' in major_str or 'HEALTH' in major_str:
        return "Public Health/Epi"
    
    else:
        return "Other"

students['Major_Category'] = students['Major'].apply(categorize_major)

print(f"\nMajor categories:")
print(students['Major_Category'].value_counts())

# Categorize enrollment type
def categorize_enrollment(enrollment):
    if pd.isna(enrollment):
        return "Unknown"
    enrollment_str = str(enrollment).upper()
    if 'PHD' in enrollment_str:
        return "PhD"
    elif 'MASTER' in enrollment_str:
        return "Master"
    elif 'UNDERGRADUATE' in enrollment_str:
        return "Undergraduate"
    else:
        return "Other"

students['Enrollment_Category'] = students['Enrollment'].apply(categorize_enrollment)

# Initialize section assignments
sections = {time: [] for time in section_times}

# Greedy assignment algorithm with demographic balancing
def calculate_section_diversity_score(section_students):
    """Calculate a diversity score for a section (lower is better/more diverse)"""
    if len(section_students) == 0:
        return 0
    
    # Count enrollment types
    enrollment_counts = defaultdict(int)
    major_counts = defaultdict(int)
    
    for student_idx in section_students:
        enrollment_counts[students.loc[student_idx, 'Enrollment_Category']] += 1
        major_counts[students.loc[student_idx, 'Major_Category']] += 1
    
    # Calculate variance (higher variance = less diverse)
    enrollment_variance = np.var(list(enrollment_counts.values())) if enrollment_counts else 0
    major_variance = np.var(list(major_counts.values())) if major_counts else 0
    
    return enrollment_variance + major_variance

def get_best_section_for_student(student_idx, available_sections):
    """Find the best section for a student to maintain balance"""
    best_section = None
    best_score = float('inf')
    
    for section_time in available_sections:
        # Simulate adding this student to the section
        test_section = sections[section_time] + [student_idx]
        
        # Calculate the diversity score
        diversity_score = calculate_section_diversity_score(test_section)
        
        # Prefer sections with fewer students (to balance size)
        size_penalty = len(sections[section_time]) * 10
        
        total_score = diversity_score + size_penalty
        
        if total_score < best_score:
            best_score = total_score
            best_section = section_time
    
    return best_section

# Sort students by number of available times (least flexible first)
students['Num_Available'] = students['Available_Times'].apply(len)
students_sorted = students.sort_values('Num_Available').reset_index(drop=True)

# Track assignments
assigned_students = set()
unassigned_students = []

# First pass: assign students with limited availability
for idx in students_sorted.index:
    available_times = students_sorted.loc[idx, 'Available_Times']
    
    if len(available_times) == 0:
        unassigned_students.append(idx)
        continue
    
    # Find the best section among available times
    best_section = get_best_section_for_student(idx, available_times)
    sections[best_section].append(idx)
    assigned_students.add(idx)

# Create output dataframe
output_rows = []
for section_time, student_indices in sections.items():
    print(f"\n{section_time}: {len(student_indices)} students")
    
    if len(student_indices) > 0:
        section_students = students_sorted.loc[student_indices]
        print(f"  Enrollment: {section_students['Enrollment_Category'].value_counts().to_dict()}")
        print(f"  Major Category: {section_students['Major_Category'].value_counts().to_dict()}")
        
        for idx in student_indices:
            student = students_sorted.loc[idx]
            output_rows.append({
                'Section_Time': section_time,
                'Name': student['Name'],
                'Email': student['Email'],
                'Enrollment': student['Enrollment'],
                'Major': student['Major'],
                'Class_Year': student['Class_Year'],
                'Enrollment_Category': student['Enrollment_Category'],
                'Major_Category': student['Major_Category']
            })

# Create output DataFrame
output_df = pd.DataFrame(output_rows)

# Save to Excel
output_file = 'section_assignments.xlsx'
output_df.to_excel(output_file, index=False)

print(f"\n{'='*60}")
print(f"Section assignments saved to: {output_file}")
print(f"Total students assigned: {len(output_rows)}")
print(f"Unassigned students: {len(unassigned_students)}")

# Print summary statistics
print(f"\n{'='*60}")
print("OVERALL SUMMARY")
print(f"{'='*60}")
for section_time in section_times:
    students_in_section = output_df[output_df['Section_Time'] == section_time]
    print(f"\n{section_time}")
    print(f"  Total: {len(students_in_section)}")
    print(f"  Enrollment: {students_in_section['Enrollment_Category'].value_counts().to_dict()}")
    print(f"  Major Category: {students_in_section['Major_Category'].value_counts().to_dict()}")

# Additional summary by demographics
print(f"\n{'='*60}")
print("DEMOGRAPHIC BALANCE CHECK")
print(f"{'='*60}")
print("\nEnrollment distribution across sections:")
enrollment_pivot = pd.crosstab(output_df['Section_Time'], output_df['Enrollment_Category'])
print(enrollment_pivot)
print("\nMajor category distribution across sections:")
major_pivot = pd.crosstab(output_df['Section_Time'], output_df['Major_Category'])
print(major_pivot)