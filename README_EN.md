# Excel Ordering Tool (GUI Application)

## Description

This tool automates order processing from a specifically structured Excel file. It enables fast and accurate calculation of the number of products to reorder based on current stock levels, product type, and business logic. The application includes a clear GUI interface and supports two calculation modes – “order coverage” and “stock replenishment”.

---

## Features

- Two calculation modes:
  - Order coverage (based on outstanding orders)
  - Stock replenishment using custom coefficients
- Logic includes:
  - Bestsellers
  - Clearance items
  - Custom stock coefficients
- Automatic data highlighting in the output (colors and borders) for better visual orientation
- Dynamic calculation of recommended order quantity based on set logic
- User-defined parameter entry
- Intuitive GUI interface for non-technical users
- Option to export the modified Excel file

---

## Usage

1. Run the application:
   ```bash
   python ordering_app.py
   ```

2. Choose a mode:
   - **Order coverage** (optionally include bestsellers)
   - **Stock replenishment** (set coefficients for bestsellers and regular items)

3. Select the input `.xlsx` file with structured data.

4. After processing, you’ll be prompted to save the output Excel file.

> Note: The application is designed for a specific data format tailored to a particular e-commerce environment. It is not a universal solution for all types of input files.

---

## Dependencies

- `pandas`
- `openpyxl`
- `tkinter` (comes with Python on Windows)

Recommended installation:
```bash
pip install -r requirements.txt
```

---

## Author & License Note

This application was fully designed, developed with the help of AI, and tested by **Tadeáš Galbavý**.  
It is a fully functional automation tool used in production for an e-commerce business operating in 6 countries.

The project showcases real-world automation for e-commerce processes and is licensed under **CC BY-NC 4.0** (non-commercial use only).
