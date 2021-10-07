# Simple XLSX splitter

## Setup

### ENV
```
virtualenv .env
. .env/bin/activate
pip install -r requirements.txt
python run.py -h
....
deactivate
```

### FILE

Store your file somewhere on your hard drive

Investigate it to identify groupper column and its name

### Launch

```
. .env/bin/activate
mkdir target
python run.py -i /path/to/your/file.xlsx -f myfield -o target
deactivate
```

Enjoy!

## License

MIT LICENSE WITHOUT WARRANTY OF ANY KIND
