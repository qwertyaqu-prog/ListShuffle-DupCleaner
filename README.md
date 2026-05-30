# ListShuffle-DupCleaner 🛠️

A simple, lightweight desktop application built with Python and Tkinter to manage text lists. This tool provides features to split lists into Excel files (with optional shuffling) and remove duplicates or filter data using comparative lists.

## 🚀 Features

### 1. Pembagi List (List Splitter & Shuffler)
* **Real-time Counter:** Counts total list items instantly.
* **Two Splitting Modes:** * **Jumlah Kolom:** Divide list into a specific number of columns.
    * **Item per Kolom:** Specify how many items should be in each column.
* **Shuffle Option:** Randomize the list before splitting.
* **Excel Export:** Generates formatted `.xlsx` spreadsheet dynamically using `openpyxl`.

### 2. Duplikat List (Duplicate Cleaner & Comparator)
* **Deduplication:** Find and remove duplicate entries instantly while automatically sorting the clean output.
* **List Comparison:** Compare a main list (Data Utama) against a benchmark list (Data Pembanding) and subtract matching elements.
* **Visual Logs:** Separate view windows for clean results and removed items.

---

## 🛠️ Requirements

* Python 3.x
* `openpyxl` library (for Excel export manipulation)

---

## ⚙️ Installation & Usage

1. Clone this repository:
   ```bash
   git clone [https://github.com/qwertyaqu-prog/ListShuffle-DupCleaner.git](https://github.com/qwertyaqu-prog/ListShuffle-DupCleaner.git)