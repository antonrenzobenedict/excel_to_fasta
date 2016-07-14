virtualenv .venv
. .venv/bin/activate
pip install openpyxl
python excel_to_fasta.py $1 $2
