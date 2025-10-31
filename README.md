📊 Excel Grade Writer with openpyxl
This Python project generates a formatted Excel spreadsheet containing student grades across multiple subjects using the openpyxl library. It dynamically calculates average scores per subject and applies styling for readability.

🚀 Features
- Creates a new Excel workbook with a sheet titled Grades
- Populates student names and their scores in Math, Science, English, and Gym
- Automatically calculates average scores for each subject
- Applies bold and colored headers for better visual presentation
- Saves the final output as NewF1_Grades.xlsx
  
📦 Requirements
- Python 3.x
- openpyxl
Install with:
pip install openpyxl


To install the required Python package, run the following command in your terminal:

```bash
pip install openpyxl
```


🧠 How It Works
- Data Setup: A dictionary holds student names and their subject scores.
- Workbook Creation: A new Excel workbook is initialized and the active sheet is renamed to "Grades".
- Header Row: The first row contains column headers like "Name", "Math", "Science", etc.
- Data Insertion: Each student's grades are inserted row by row.
- Average Calculation: A formula is added to row 7 to compute the average score for each subject.
- Styling: Header cells are styled with bold font and a blue tint.
- Save File: The workbook is saved as NewF1_Grades.xlsx.

- 
## 🖨️ Output Preview

The script creates an Excel file named `NewF1_Grades.xlsx` with the following structure:

| Name    | Math | Science | English | Gym  |
|---------|------|---------|---------|------|
| Max     | 65   | 78      | 98      | 89   |
| Charles | 55   | 72      | 87      | 95   |
| Lando   | 100  | 45      | 75      | 92   |
| Oscar   | 30   | 25      | 45      | 100  |
| Carlos  | 100  | 100     | 100     | 60   |
| **Avg** | 70.0 | 64.0    | 81.0    | 87.2 |

- The **header row** is styled with bold font and a blue tint.
- The **last row** contains Excel formulas that calculate the average score for each subject.
- The file is saved in the same directory as `NewF1_Grades.xlsx`.

> You can open the file in Excel to view the formulas and formatting in action. 

🛠️ Customization
You can easily:
- Add more students or subjects
- Change the styling (font, color, alignment)
- Extend with charts or conditional formatting
📚 References
- [openpyxl documentation](https://openpyxl.readthedocs.io/en/stable/)


