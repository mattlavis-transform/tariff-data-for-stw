# Generate misc data extracts for Single Trade Window

## Implementation steps

- Create and activate a virtual environment, e.g.

  `python3 -m venv venv/`
  `source venv/bin/activate`

- Install necessary Python modules 

  - autopep8==1.5.5
  - psycopg2==2.8.6
  - pycodestyle==2.6.0
  - python-dotenv==0.15.0
  - toml==0.10.2
  - XlsxWriter==1.3.7

  via `pip3 install -r requirements.txt`

## Usage

### To create the extract
`python3 create.py`
