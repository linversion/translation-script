a python script help you parse the multi language excel into specific format.
--
```python
# change the format by your needs
formatter = '<string name=\"{name}\">{str}</string>\n'
values_folders = ['values', 'values-zh-rCN']
# your project's res folder path
res_path = './res/'
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
        end_col=3,
        export_direct_to_res=False
    )
```

step 2: checkout the result in result.xml

![excel](https://raw.githubusercontent.com/linversion/translation-script/main/excel.png)

![result](https://raw.githubusercontent.com/linversion/translation-script/main/result.png)

### optional
if you set the export_direct_to_res to True, it will directly add the new translation to your strings.xml.If there are a same key already, it will simply update the string text.And string array will remain the same of course.