"""
call_bin.py
"""
import subprocess
from pyhandytools.file import FileUtil

excel_data = [{...}]
sheet_name = 'xxx'
lst_json = './data/tmp_lst.json'
FileUtils.write2json(lst_json, excel_data)
# 定义命令和参数
command = [
    "./bin/tools/lst2xlsx",
    "-jp", lst_json,
    "-sp", excel_path,
    "-s", sheet_name,
    "-ord", json.dumps(TASK_EVAL_COLUMN_ORDER, ensure_ascii=False)
]
try:
    result = subprocess.run(command, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    print("OK!", result)
except subprocess.CalledProcessError as e:
    print("Error:", e.stderr)