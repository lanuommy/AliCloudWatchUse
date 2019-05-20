import xlwt
import json
from aliyunsdkcore.client import AcsClient
from aliyunsdkcore.request import CommonRequest

#创建excel
book = xlwt.Workbook(encoding='utf-8',style_compression=0)
#添加sheet
sheet = book.add_sheet('阿里云监控报告2019年4月',cell_overwrite_ok=True)
#配置阿里云的ak
client = AcsClient('ak', 'ak', 'cn-beijing')
#写入表头，定义好列
sheet.write(0, 0, label = 'InstanceID')
sheet.write(0, 1, label = 'HostName')
sheet.write(0, 2, label = 'CPU最大')
sheet.write(0, 3, label = 'CPU最小')
sheet.write(0, 4, label = 'CPU平均')
sheet.write(0, 5, label = '内存最大')
sheet.write(0, 6, label = '内存最小')
sheet.write(0, 7, label = '内存平均')
sheet.write(0, 10, label = '没有被检测出来的实例')
#定义QUeryMetricsList的api调用
def aliCheckData(project, instanceID):
   a = '[{"instanceId":"' + instanceID + '"}]'
   #   print(a)
   request = CommonRequest()
   request.set_accept_format('json')
   request.set_domain('metrics.cn-hangzhou.aliyuncs.com')
   request.set_method('POST')
   request.set_protocol_type('https')  # https | http
   request.set_version('2018-03-08')
   request.set_action_name('QueryMetricList')
   request.add_query_param('RegionId', 'cn-beijing')
   request.add_query_param('Metric', project)
   request.add_query_param('Period', '7200')
   request.add_query_param('StartTime', '2019-04-01 00:00:00')
   request.add_query_param('EndTime', '2019-04-30 00:00:00')
   request.add_query_param('Project', 'acs_ecs_dashboard')
   # request.add_query_param('Dimensions', '[{"instanceId": instanceID}]')
   request.add_query_param('Dimensions', a)
   response = client.do_action(request)
   #print(response)
   cpu_total = eval(response)['Datapoints']
   return cpu_total

#定义DescribeInstances的所有实例信息获取
def aliInstancesInfo():
   request = CommonRequest()
   #request = CommonRequest()
   request.set_accept_format('json')
   request.set_domain('ecs.aliyuncs.com')
   request.set_method('POST')
   request.set_protocol_type('https') # https | http
   request.set_version('2014-05-26')
   request.set_action_name('DescribeInstances')

   request.add_query_param('RegionId', 'cn-beijing')
   request.add_query_param('PageNumber', '1')
   request.add_query_param('PageSize', '100')
   response = client.do_action(request)
   response = (str(response, encoding='utf-8'))
   data = json.loads(response)
   #InfoList = []
   #定义n为从excel的第二行开始
   n = 1
   #大循环，遍历所有instanceid，并定义计算的初始值，平均值和最小值为计算出来的结果，最大值为取列表最大
   for i in data["Instances"]["Instance"]:
#   for i in "p":

      #从api中获取第一个循环里的ID与CPU的值的获取
      cpu_total = aliCheckData('cpu_total', i["InstanceId"])
      ##从api中获取第一个循环里的ID与Memory的值的获取
      memory_used = aliCheckData('memory_usedutilization', i["InstanceId"])
      #      memory_used = aliCheckData('memory_usedutilization', i["InstanceId"])
      #如果找到云监控中没有数值的信息，则跳过
      if (len(eval(cpu_total)) == 0):
         print("len:0", i["InstanceId"])
         sheet.write(n, 10, i["InstanceId"])
         continue
      if (len(eval(memory_used)) == 0):
         print("len:0", i["InstanceId"])
         continue
      #a = i["InstanceId"]
      #从QueryMetricsList中获取InstanceID与Hostname两个值
      instanceID =  i["InstanceId"]
      hostname = i["InstanceName"]
      print(hostname)
      print(instanceID)
      #写入表格
      sheet.write(n, 0, instanceID)
      sheet.write(n, 1, hostname)
      #定义一个列表，为了后面循环中持续向列表中写值，方便比较大小
      cpu_total_max = []
      #开始循环取值
      cpu_sum_average = 0
      cpu_sum_max = 0
      cpu_sum_min = 0
      for i in range(len(eval(cpu_total))):

         # sum_list.append(eval(cpu_total)[i]['Average'])
         cpu_sum_average = cpu_sum_average + eval(cpu_total)[i]['Average']
         #print(eval(cpu_total)[i]['Average'])
         cpu_sum_max = cpu_sum_max + eval(cpu_total)[i]['Maximum']
         #cpu_total_max.append(eval(cpu_total)[i]['Maximum'])
         cpu_sum_min = cpu_sum_min + eval(cpu_total)[i]['Minimum']

      print('cpu最大值为：',str(cpu_sum_max))
      print('cpu最小值为：',str(cpu_sum_min))
      print('cpu平均值为：',str(cpu_sum_average))
      print('cpu信息个数为：',len(eval(cpu_total)))
      print(len(eval(cpu_total)))
      cpu_average_res = str(cpu_sum_average / len(eval(cpu_total)))
      cpu_max_res = str(cpu_sum_max / len(eval(cpu_total)))
      #cpu_max_res = (max(cpu_total_max))
      cpu_min_res = str(cpu_sum_min / len(eval(cpu_total)))
      print(cpu_max_res)
      print(cpu_min_res)
      print(cpu_average_res)
      sheet.write(n, 2, cpu_max_res)
      sheet.write(n, 3, cpu_min_res)
      sheet.write(n, 4, cpu_average_res)
      #定义memory的取值列表，方便后面取大小
      #memory_used_max = []
      mem_sum_average = 0
      mem_sum_min = 0
      mem_sum_max = 0
      for i in range(len(eval(memory_used))):
         # sum_list.append(eval(cpu_total)[i]['Average'])
         mem_sum_average = mem_sum_average + eval(memory_used)[i]['Average']
         mem_sum_max = mem_sum_max + eval(memory_used)[i]['Maximum']
         #memory_used_max.append(eval(memory_used)[i]['Maximum'])
         mem_sum_min = mem_sum_min + eval(memory_used)[i]['Minimum']

      print('内存最大值为：',str(mem_sum_max))
      print('内存最小值为：',str(mem_sum_min))
      print('内存平均值为：',str(mem_sum_average))
      print('内存信息个数为：',len(eval(memory_used)))

      mem_average_res = str(mem_sum_average / len(eval(memory_used)))
      mem_max_res = str(mem_sum_max / len(eval(memory_used)))
      #mem_max_res = max(memory_used_max)
      mem_min_res = str(mem_sum_min / len(eval(memory_used)))
      print(mem_max_res)
      print(mem_min_res)
      print(mem_average_res)
      sheet.write(n, 5, mem_max_res)
      sheet.write(n, 6, mem_min_res)
      sheet.write(n, 7, mem_average_res)

      #每次的行数递增，这样可以循环往下写
      n += 1
aliInstancesInfo()

book.save('/Users/qiao/阿里云监控报告2019年4月.xls')

