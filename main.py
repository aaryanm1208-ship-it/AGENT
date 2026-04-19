from excel_reader import load_excel
import os

print("Excel Agent Started (type 'exit' to quit)\n")

while True:
    file_input = input("Enter Excel file name: ").lower()

    if file_input == "exit":
        break

    # -------- AUTO FILE DETECT (no need .xlsx) --------
    if not file_input.endswith(".xlsx"):
        file_input += ".xlsx"

    if not os.path.exists(file_input):
        print("File not found\n")
        continue

    df = load_excel(file_input)

    if df is None:
        print("Error loading file\n")
        continue

    print(f"\nLoaded {file_input}. Ask questions (type 'back' to change file)\n")

    while True:
        query = input("Ask: ").lower()

        if query == "back":
            break

        if query == "exit":
            exit()

        found = False

        # -------- COLUMN BASED (natural language) --------
        for col in df.columns:
            col_lower = col.lower()

            if col_lower in query:
                found = True

                if "total" in query or "sum" in query:
                    print(f"\nTotal {col} =", df[col].sum())

                elif "average" in query or "mean" in query:
                    print(f"\nAverage {col} =", df[col].mean())

                elif "max" in query or "highest" in query:
                    print(f"\nMax {col} =", df[col].max())

                elif "min" in query or "lowest" in query:
                    print(f"\nMin {col} =", df[col].min())

                elif "count" in query:
                    print(f"\nCount {col} =", df[col].count())

                # -------- TOP 5 --------
                elif "top" in query:
                    print(f"\nTop 5 {col}:\n", df.nlargest(5, col))

                break

        # -------- FILTER (where condition) --------
        if "where" in query:
            try:
                parts = query.split("where")[1].strip().split("=")
                column = parts[0].strip()
                value = parts[1].strip()

                result = df[df[column].astype(str).str.lower() == value]

                if not result.empty:
                    print("\nFiltered Data:\n", result.head(10))
                    found = True
                else:
                    print("\nNo matching data found")
                    found = True
            except:
                print("\nInvalid filter format. Use: sales where region = west")
                found = True

        # -------- VALUE SEARCH --------
        if not found:
            for col in df.columns:
                try:
                    result = df[df[col].astype(str).str.lower().str.contains(query)]
                    if not result.empty:
                        print("\nMatching Rows:\n", result.head(10))
                        found = True
                        break
                except:
                    continue

        # -------- FALLBACK --------
        if not found:
            print("\nTry:")
            print("- total sales")
            print("- highest runs")
            print("- top runs")
            print("- sales where region = west")

        print("-" * 50)
