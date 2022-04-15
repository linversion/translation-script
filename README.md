```python
# change the format by your needs
formatter = '<string name=\"{name}\">{str}</string>\n'

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    reload(sys)
    sys.setdefaultencoding('utf8')
    # step 1: your input
    read_excel(
        'trans-sample.xlsx',
        start_column=2,
        start_row=2,
        key_column=1,
        end_row=5,
        end_col=3
    )
```

step 2: checkout the result in result.xml
