# nvdmonitor
**nvdmonitor可以定时监控NVD漏洞库指定目标，并通过企业微信定时推送新增漏洞详情**<br>
pip install -r requirements.txt<br>
##config.yaml配置扫描的参数<br>
##单次扫描本日新增漏洞并推送到企微<br>
python nvdmonitor.py   
##每日定时扫描本日新增漏洞并推送到企微<br>
python scheduler.py     

<br>

## 通过config.yaml修改扫描参数<br>
#### 1.产品名称配置(NVD搜索的内容)<br>
querry: <br>
#### 2.扫描页数<br>
page: 1-5<br>
#### 3.任务名称<br>
target: <br>

#### 4.企微机器人推送配置(为空则不推送)<br>
wechat_bot:<br>
  name: nvdmonitor<br>
  webhook_url: " "<br>

#### 5.调度配置
scheduler:
  mode: "fixed"  # fixed 或 interval
  fixed模式配置 (每天固定时间执行)<br>
  fixed_times:<br>
    - "09:00"<br>
    - "14:00"<br>
    - "18:00"<br>
  interval模式配置 (每隔几小时执行)<br>
  interval_hours: 6<br>
<br>

<font color="red">注意：企微推送有字数限制，如果消息过长则不能推送；通过任务日期json去重；扫描结果会存储在result</font>

