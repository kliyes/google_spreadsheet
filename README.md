=====
Google Spreadsheets API Python Client
=====

对[google-api-python-client](https://pypi.python.org/pypi/google-api-python-client/)进一步封装,使其更加便捷的操作spreadsheet文档

## Installation

### 使用pip安装(推荐)
```bash
$ pip install git+ssh://git@github.stm.com/jingyang/google_spreadsheet.git
```

### 源码安装
```bash
$ git clone git@github.stm.com:jingyang/google_spreadsheet.git
$ cd google_spreadsheet
$ python setup.py install
```

### requirements.pip文件
```requirements.pip
luigi
gsutil
git+ssh://git@github.stm.com/jingyang/google_spreadsheet.git
...
```

## Basic Usage

### 获取OAuth2认证key文件
1. 登录console.cloud.google.com
2. 选择现有的或新建一个project
3. 进入API Manager > Library,搜索Google Sheet API,启用;
4. 进入IAM & Admin > Service accounts,新建一个service account并下载对应的json格式key文件
   (若在Google Compute Engine(GCE)上执行程序,可以跳过这一步)
5. 将要操作的spreadsheet文档分享给刚才创建的service account的Service account ID(service-account-name@project-id.iam.gserviceaccount.com)
   (如果使用GCE,则分享给Compute Engine default service account)

### 认证授权
```python
from google_spreadsheet import spreadsheet_service

service = spreadsheet_service(key_file_location=/path/to/your/service_account_key_file.json)
```
如果是在Google Compute Engine(GCE)上运行的程序:
```python
service = spreadsheet_service(on_gce=True)
```

### 基本操作
```python
from google_spreadsheet import Client

client = Client(service)

# file id可以从文档的URL中获取
spreadsheet = client.open("1nQiHsUMCZB-pa8MRurkzvsbesmPWevDbSWj7VIhwPcU")
print spreadsheet.title
```