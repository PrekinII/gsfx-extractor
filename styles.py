from openpyxl.styles import Font, PatternFill, Alignment


HEADER_FILL = PatternFill(start_color='B0C4DE', end_color='B0C4DE', fill_type='solid')
HEADER_FONT = Font(bold=True)
ALIGN_RIGHT = Alignment(horizontal='right')
ALIGN_CENTER = Alignment(horizontal='center')
NUM_FORMAT = '#,##0.00'