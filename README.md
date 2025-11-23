
# Dynamic Sampling VBA Macro

## Overview
This VBA macro dynamically creates a sample from a large dataset in Excel while ensuring:
- The sample size does not exceed **999 rows**.
- Sampling is spread across the entire dataset (top to bottom).
- Handles edge cases where unique categories are very large.
- Preserves original logic by adjusting step size dynamically.

## Features
- **Dynamic Step Size Adjustment**: Automatically increases step size until sample size ≤ 999.
- **Category Awareness**: Builds a dictionary of unique categories for potential proportional sampling.
- **No Manual Reruns**: The macro self-adjusts for any dataset size.

## How It Works
1. User specifies the column containing categories.
2. The macro calculates the total rows and unique categories.
3. It iteratively adjusts the step size until the sample size is within the limit.
4. Colors sampled rows and deletes non-sampled rows.

## Usage
1. Open Excel and press `Alt + F11` to open the VBA editor.
2. Insert a new module and paste the code from `RandomizeDynamicSample.bas`.
3. Run the macro `RandomizeDynamicSample`.
4. Enter the column letter when prompted.

## Example
- Total rows: 13,429
- Unique categories: 3,022
- Max sample: 999 rows
- Final step size: dynamically calculated (e.g., 14)

## License
Free to use and modify.
