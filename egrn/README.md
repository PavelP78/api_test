# Участок 47:01:1516001:2322 (выписка ЕГРН / ЕПГУ)

Разбор выписки: [`ANALIZ.md`](ANALIZ.md)

```bash
python3 -m pip install -r egrn/requirements.txt
python3 egrn/generate_schema.py
python3 -m pytest egrn/test_geometry.py -q
```
