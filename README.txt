GEORGE WASHINGTON SCRIPTS

There are two versioned script folders. Each block below can be copied and pasted into Terminal.

Folder layout:

```text
george_washington/
├── README.txt
├── .gitignore
├── .venv/
├── george_washington_v1/
│   ├── george_washington_v1.py
│   ├── george_washington_v1_input.xlsx
│   ├── george_washington_v1_scatterplot.png
│   ├── george_washington_v1_with_data.xlsx
│   └── george_washington_v1_3d_plot.html
└── george_washington_v2/
    ├── george_washington_v2.py
    ├── george_washington_v2_input.xlsx
    ├── george_washington_v2_scatterplot.png
    └── george_washington_v2_with_data.xlsx
```


BLOCK 1: george_washington_v1

Folder:
george_washington_v1/

Script:
george_washington_v1.py

Input:
george_washington_v1_input.xlsx

Outputs:
george_washington_v1_scatterplot.png
george_washington_v1_with_data.xlsx
george_washington_v1_3d_plot.html

Copy and paste this block into Terminal:

```bash
cd george_washington/george_washington_v1
python3 -m venv ../.venv
source ../.venv/bin/activate
python3 -m pip install --upgrade pip
python3 -m pip install pandas matplotlib openpyxl plotly
python3 george_washington_v1.py
deactivate
```


BLOCK 2: george_washington_v2

Folder:
george_washington_v2/

Script:
george_washington_v2.py

Input:
george_washington_v2_input.xlsx

Outputs:
george_washington_v2_scatterplot.png
george_washington_v2_with_data.xlsx

Copy and paste this block into Terminal:

```bash
cd george_washington/george_washington_v2
python3 -m venv ../.venv
source ../.venv/bin/activate
python3 -m pip install --upgrade pip
python3 -m pip install pandas matplotlib openpyxl plotly
python3 george_washington_v2.py
deactivate
```


OPTIONAL: Open plot windows too

For Block 1, replace this line:

```bash
python3 george_washington_v1.py
```

with:

```bash
python3 george_washington_v1.py --show
```

For Block 2, replace this line:

```bash
python3 george_washington_v2.py
```

with:

```bash
python3 george_washington_v2.py --show
```
