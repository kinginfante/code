Attribute VB_Name = "ModBx"
Option Explicit


Public Sub dtgKj(Lb As Integer)
frmFYBX.dtgBx.Columns("日期").Visible = True
frmFYBX.dtgBx.Columns("福利费").Visible = False
frmFYBX.dtgBx.Columns("房屋补贴").Visible = False
frmFYBX.dtgBx.Columns("旅游费").Visible = False
frmFYBX.dtgBx.Columns("高温费").Visible = False
frmFYBX.dtgBx.Columns("通信费").Visible = False
frmFYBX.dtgBx.Columns("市内交通费").Visible = False
frmFYBX.dtgBx.Columns("市外交通费").Visible = False
frmFYBX.dtgBx.Columns("运费").Visible = False
frmFYBX.dtgBx.Columns("住宿费").Visible = False
frmFYBX.dtgBx.Columns("部门团队费").Visible = False
frmFYBX.dtgBx.Columns("餐费").Visible = False
frmFYBX.dtgBx.Columns("招待费").Visible = False
frmFYBX.dtgBx.Columns("礼品费").Visible = False
frmFYBX.dtgBx.Columns("房租").Visible = False
frmFYBX.dtgBx.Columns("物业费").Visible = False
frmFYBX.dtgBx.Columns("水电").Visible = False
frmFYBX.dtgBx.Columns("电话").Visible = False
frmFYBX.dtgBx.Columns("办公用品").Visible = False
'frmFYBX.dtgBx.Columns("邮资").Visible = False
frmFYBX.dtgBx.Columns("市场推广").Visible = False
frmFYBX.dtgBx.Columns("人员招聘").Visible = False
frmFYBX.dtgBx.Columns("快递费").Visible = False
frmFYBX.dtgBx.Columns("培训费").Visible = False
frmFYBX.dtgBx.Columns("财务手续费").Visible = False
frmFYBX.dtgBx.Columns("团队建设费").Visible = False
frmFYBX.dtgBx.Columns("停车费").Visible = False
frmFYBX.dtgBx.Columns("车辆费").Visible = False
frmFYBX.dtgBx.Columns("公共停车费").Visible = False
frmFYBX.dtgBx.Columns("公共车辆费").Visible = False
frmFYBX.dtgBx.Columns("工具").Visible = False
frmFYBX.dtgBx.Columns("易耗").Visible = False
frmFYBX.dtgBx.Columns("外劳").Visible = False
frmFYBX.dtgBx.Columns("交通补贴").Visible = False
frmFYBX.dtgBx.Columns("驻外津贴").Visible = False
frmFYBX.dtgBx.Columns("岗位补贴").Visible = False
frmFYBX.dtgBx.Columns("综合保险").Visible = False
frmFYBX.dtgBx.Columns("三金").Visible = False
frmFYBX.dtgBx.Columns("公积金").Visible = False
frmFYBX.dtgBx.Columns("合同编号").Visible = False
frmFYBX.dtgBx.Columns("部门").Visible = False
'frmFYBX.dtgBx.Columns("福利").Visible = False
frmFYBX.dtgBx.Columns("区域").Visible = False
frmFYBX.dtgBx.Columns("归属人").Visible = False
frmFYBX.dtgBx.Columns("归属人签字").Visible = False
frmFYBX.dtgBx.Columns("签字时间").Visible = False
frmFYBX.dtgBx.Columns("福利费").Visible = False
frmFYBX.dtgBx.Columns("部门经理签字").Visible = False
frmFYBX.dtgBx.Columns("签字日期").Visible = False
frmFYBX.dtgBx.Columns("出租车注明").Visible = False
frmFYBX.dtgBx.Columns("签收日期").Visible = False
frmFYBX.frmRen.Visible = False
frmFYBX.dtgNx.Visible = True
frmFYBX.dtgBx.Visible = False
Select Case Lb
Case 7 '公共费用
    frmFYBX.dtgBx.Columns("房租").Visible = True
    'frmFYBX.dtgBx.Columns("物业费").Visible = True
    frmFYBX.dtgBx.Columns("水电").Visible = True
    frmFYBX.dtgBx.Columns("电话").Visible = True
    frmFYBX.dtgBx.Columns("办公用品").Visible = True
    'frmFYBX.dtgBx.Columns("邮资").Visible = True
    frmFYBX.dtgBx.Columns("市场推广").Visible = True
    frmFYBX.dtgBx.Columns("人员招聘").Visible = True
    frmFYBX.dtgBx.Columns("快递费").Visible = True
    frmFYBX.dtgBx.Columns("培训费").Visible = True
    frmFYBX.dtgBx.Columns("财务手续费").Visible = True
    frmFYBX.dtgBx.Columns("福利费").Visible = True
    frmFYBX.dtgBx.Columns("公共停车费").Visible = True
    frmFYBX.dtgBx.Columns("公共车辆费").Visible = True
Case 8 '总经理室
    frmFYBX.dtgBx.Columns("通信费").Visible = True
    frmFYBX.dtgBx.Columns("市内交通费").Visible = True
    frmFYBX.dtgBx.Columns("市外交通费").Visible = True
    frmFYBX.dtgBx.Columns("住宿费").Visible = True
    frmFYBX.dtgBx.Columns("餐费").Visible = True
    frmFYBX.dtgBx.Columns("招待费").Visible = True
    frmFYBX.dtgBx.Columns("礼品费").Visible = True
    frmFYBX.dtgBx.Columns("车辆费").Visible = True
Case 50 '运费
    frmFYBX.dtgBx.Columns("运费").Visible = True
    frmFYBX.dtgBx.Columns("合同编号").Visible = True
    frmFYBX.dtgBx.Columns("部门").Visible = True
    frmFYBX.dtgBx.Columns("区域").Visible = True
    frmFYBX.dtgBx.Columns("归属人").Visible = True
    frmFYBX.dtgBx.Columns("归属人签字").Visible = True
    frmFYBX.dtgBx.Columns("签字时间").Visible = True
    frmFYBX.dtgBx.Columns("部门经理签字").Visible = True
    frmFYBX.dtgBx.Columns("签字日期").Visible = True

Case 51 '运费
    frmFYBX.dtgBx.Columns("运费").Visible = True
    frmFYBX.dtgBx.Columns("合同编号").Visible = True
    frmFYBX.dtgBx.Columns("部门").Visible = True
    frmFYBX.dtgBx.Columns("区域").Visible = True
    frmFYBX.dtgBx.Columns("归属人").Visible = True
    frmFYBX.dtgBx.Columns("归属人签字").Visible = True
    frmFYBX.dtgBx.Columns("签字时间").Visible = True
    frmFYBX.dtgBx.Columns("部门经理签字").Visible = True
    frmFYBX.dtgBx.Columns("签字日期").Visible = True

Case 10 '福利

Case 11 '工程外地
    frmFYBX.dtgBx.Columns("市外交通费").Visible = True
    frmFYBX.dtgBx.Columns("住宿费").Visible = True
    frmFYBX.dtgBx.Columns("餐费").Visible = True
    frmFYBX.dtgBx.Columns("合同编号").Visible = True
    frmFYBX.dtgBx.Columns("部门").Visible = True
    frmFYBX.dtgBx.Columns("区域").Visible = True
    frmFYBX.dtgBx.Columns("归属人").Visible = True
    frmFYBX.dtgBx.Columns("归属人签字").Visible = True
    frmFYBX.dtgBx.Columns("签字时间").Visible = True
    frmFYBX.dtgBx.Columns("部门经理签字").Visible = True
    frmFYBX.dtgBx.Columns("签字日期").Visible = True
    'frmFYBX.dtgBx.Columns("工具费").Visible = True
    frmFYBX.dtgBx.Columns("易耗").Visible = True
    frmFYBX.dtgBx.Columns("外劳").Visible = True
Case 12 '工程外地
    frmFYBX.dtgBx.Columns("市外交通费").Visible = True
    frmFYBX.dtgBx.Columns("住宿费").Visible = True
    frmFYBX.dtgBx.Columns("餐费").Visible = True
    frmFYBX.dtgBx.Columns("合同编号").Visible = True
    frmFYBX.dtgBx.Columns("部门").Visible = True
    frmFYBX.dtgBx.Columns("区域").Visible = True
    frmFYBX.dtgBx.Columns("归属人").Visible = True
    frmFYBX.dtgBx.Columns("归属人签字").Visible = True
    frmFYBX.dtgBx.Columns("签字时间").Visible = True
    frmFYBX.dtgBx.Columns("部门经理签字").Visible = True
    frmFYBX.dtgBx.Columns("签字日期").Visible = True
    'frmFYBX.dtgBx.Columns("工具费").Visible = True
    frmFYBX.dtgBx.Columns("易耗").Visible = True
    frmFYBX.dtgBx.Columns("外劳").Visible = True

Case 53 '销售经理
    frmFYBX.dtgBx.Columns("通信费").Visible = True
    frmFYBX.dtgBx.Columns("市内交通费").Visible = True
    frmFYBX.dtgBx.Columns("市外交通费").Visible = True
    frmFYBX.dtgBx.Columns("住宿费").Visible = True
    frmFYBX.dtgBx.Columns("餐费").Visible = True
    frmFYBX.dtgBx.Columns("招待费").Visible = True
    frmFYBX.dtgBx.Columns("礼品费").Visible = True
    frmFYBX.dtgBx.Columns("车辆费").Visible = True
    frmFYBX.dtgBx.Columns("快递费").Visible = True
    frmFYBX.dtgBx.Columns("部门团队费").Visible = True
    frmFYBX.dtgBx.Columns("办公用品").Visible = True
    frmFYBX.dtgBx.Columns("培训费").Visible = True
    frmFYBX.dtgBx.Columns("福利费").Visible = True
    
Case 14 '部门经理
    frmFYBX.dtgBx.Columns("通信费").Visible = True
    frmFYBX.dtgBx.Columns("市内交通费").Visible = True
    frmFYBX.dtgBx.Columns("市外交通费").Visible = True
    frmFYBX.dtgBx.Columns("住宿费").Visible = True
    frmFYBX.dtgBx.Columns("餐费").Visible = True
    frmFYBX.dtgBx.Columns("招待费").Visible = True
    frmFYBX.dtgBx.Columns("礼品费").Visible = True
    frmFYBX.dtgBx.Columns("车辆费").Visible = True
    frmFYBX.dtgBx.Columns("部门团队费").Visible = True
    
Case 15 '业务员
    frmFYBX.dtgBx.Columns("通信费").Visible = True
    frmFYBX.dtgBx.Columns("市内交通费").Visible = True
    frmFYBX.dtgBx.Columns("市外交通费").Visible = True
    frmFYBX.dtgBx.Columns("住宿费").Visible = True
    frmFYBX.dtgBx.Columns("餐费").Visible = True
    frmFYBX.dtgBx.Columns("招待费").Visible = True
    frmFYBX.dtgBx.Columns("礼品费").Visible = True
    'frmFYBX.dtgBx.Columns("快递费").Visible = True
frmFYBX.dtgBx.Columns("出租车注明").Visible = True

Case 16 '业务员
    'frmFYBX.dtgBx.Columns("通信费").Visible = True
    frmFYBX.dtgBx.Columns("市内交通费").Visible = True
    frmFYBX.dtgBx.Columns("市外交通费").Visible = True
    frmFYBX.dtgBx.Columns("住宿费").Visible = True
    frmFYBX.dtgBx.Columns("餐费").Visible = True
    frmFYBX.dtgBx.Columns("招待费").Visible = True
    frmFYBX.dtgBx.Columns("礼品费").Visible = True
    'frmFYBX.dtgBx.Columns("快递费").Visible = True
frmFYBX.dtgBx.Columns("出租车注明").Visible = True

Case 17  '普通报销

   ' frmFYBX.dtgBx.Columns("通信费").Visible = True
    frmFYBX.dtgBx.Columns("市内交通费").Visible = True
    frmFYBX.dtgBx.Columns("市外交通费").Visible = True
    frmFYBX.dtgBx.Columns("住宿费").Visible = True
    frmFYBX.dtgBx.Columns("餐费").Visible = True
    frmFYBX.dtgBx.Columns("房屋补贴").Visible = True
    frmFYBX.dtgBx.Columns("办公用品").Visible = True
Case 18 '普通报销
    frmFYBX.dtgBx.Columns("通信费").Visible = True
    frmFYBX.dtgBx.Columns("市内交通费").Visible = True
    frmFYBX.dtgBx.Columns("市外交通费").Visible = True
    frmFYBX.dtgBx.Columns("住宿费").Visible = True
    frmFYBX.dtgBx.Columns("餐费").Visible = True
    frmFYBX.dtgBx.Columns("房屋补贴").Visible = True
    frmFYBX.dtgBx.Columns("办公用品").Visible = True
    
Case 20
Case 21

Case 32 '费用归属
    frmFYBX.dtgBx.Columns("通信费").Visible = True
    frmFYBX.dtgBx.Columns("市内交通费").Visible = True
    frmFYBX.dtgBx.Columns("市外交通费").Visible = True
    frmFYBX.dtgBx.Columns("住宿费").Visible = True
    frmFYBX.dtgBx.Columns("餐费").Visible = True
    frmFYBX.dtgBx.Columns("招待费").Visible = True
    frmFYBX.dtgBx.Columns("礼品费").Visible = True
    frmFYBX.dtgBx.Columns("快递费").Visible = True
    frmFYBX.dtgBx.Columns("办公用品").Visible = True
    frmFYBX.dtgBx.Columns("培训费").Visible = True
    frmFYBX.dtgBx.Columns("部门").Visible = True
    frmFYBX.dtgBx.Columns("区域").Visible = True
    frmFYBX.dtgBx.Columns("归属人").Visible = True
    frmFYBX.dtgBx.Columns("归属人签字").Visible = True
    frmFYBX.dtgBx.Columns("签字时间").Visible = True
    frmFYBX.dtgBx.Columns("部门经理签字").Visible = True
    frmFYBX.dtgBx.Columns("签字日期").Visible = True
    'frmFYBX.dtgBx.Columns("工具").Visible = True
    frmFYBX.dtgBx.Columns("福利费").Visible = True
    frmFYBX.dtgBx.Columns("易耗").Visible = True
    frmFYBX.dtgBx.Columns("外劳").Visible = True
    frmFYBX.dtgBx.Columns("车辆费").Visible = True
    'frmFYBX.frmRen.Visible = True
Case 33 '费用归属(目前已不存在)
    frmFYBX.dtgBx.Columns("通信费").Visible = True
    frmFYBX.dtgBx.Columns("市内交通费").Visible = True
    frmFYBX.dtgBx.Columns("市外交通费").Visible = True
    frmFYBX.dtgBx.Columns("住宿费").Visible = True
    frmFYBX.dtgBx.Columns("餐费").Visible = True
    frmFYBX.dtgBx.Columns("招待费").Visible = True
    frmFYBX.dtgBx.Columns("礼品费").Visible = True
    frmFYBX.dtgBx.Columns("快递费").Visible = True
    frmFYBX.dtgBx.Columns("办公用品").Visible = True
    frmFYBX.dtgBx.Columns("培训费").Visible = True
    'frmFYBX.dtgBx.Columns("工具").Visible = True
    frmFYBX.dtgBx.Columns("福利费").Visible = True
    frmFYBX.dtgBx.Columns("部门").Visible = True
    frmFYBX.dtgBx.Columns("区域").Visible = True
    frmFYBX.dtgBx.Columns("归属人").Visible = True
    frmFYBX.dtgBx.Columns("归属人签字").Visible = True
    frmFYBX.dtgBx.Columns("签字时间").Visible = True
    frmFYBX.dtgBx.Columns("部门经理签字").Visible = True
    frmFYBX.dtgBx.Columns("签字日期").Visible = True
    frmFYBX.dtgBx.Columns("易耗").Visible = True
    frmFYBX.dtgBx.Columns("外劳").Visible = True
    frmFYBX.dtgBx.Columns("车辆费").Visible = True
Case 71 '费用归属
    frmFYBX.dtgBx.Columns("通信费").Visible = True
    frmFYBX.dtgBx.Columns("市内交通费").Visible = True
    frmFYBX.dtgBx.Columns("市外交通费").Visible = True
    frmFYBX.dtgBx.Columns("住宿费").Visible = True
    frmFYBX.dtgBx.Columns("餐费").Visible = True
    frmFYBX.dtgBx.Columns("招待费").Visible = True
    frmFYBX.dtgBx.Columns("礼品费").Visible = True
    frmFYBX.dtgBx.Columns("快递费").Visible = True
    frmFYBX.dtgBx.Columns("办公用品").Visible = True
    frmFYBX.dtgBx.Columns("培训费").Visible = True
    'frmFYBX.dtgBx.Columns("工具").Visible = True
    frmFYBX.dtgBx.Columns("福利费").Visible = True
    frmFYBX.dtgBx.Columns("外劳").Visible = True
    frmFYBX.dtgBx.Columns("车辆费").Visible = True
    frmFYBX.dtgBx.Columns("部门").Visible = True
    frmFYBX.dtgBx.Columns("区域").Visible = True
    frmFYBX.dtgBx.Columns("归属人").Visible = True
    frmFYBX.dtgBx.Columns("归属人签字").Visible = True
    frmFYBX.dtgBx.Columns("签字时间").Visible = True
    frmFYBX.dtgBx.Columns("部门经理签字").Visible = True
    frmFYBX.dtgBx.Columns("签字日期").Visible = True
Case 35 '福利
    frmFYBX.dtgBx.Columns("福利费").Visible = True
    frmFYBX.dtgBx.Columns("房屋补贴").Visible = True
    frmFYBX.dtgBx.Columns("车辆费").Visible = True
    frmFYBX.dtgBx.Columns("通信费").Visible = True
    frmFYBX.dtgBx.Columns("旅游费").Visible = True
    frmFYBX.dtgBx.Columns("交通补贴").Visible = True
    frmFYBX.dtgBx.Columns("驻外津贴").Visible = True
    frmFYBX.dtgBx.Columns("岗位补贴").Visible = True
    frmFYBX.dtgBx.Columns("部门").Visible = True
    frmFYBX.dtgBx.Columns("区域").Visible = True
    frmFYBX.dtgBx.Columns("归属人").Visible = True
    frmFYBX.dtgBx.Columns("签收日期").Visible = True
    frmFYBX.dtgBx.Columns("ywyuid").Width = 0
    frmFYBX.dtgBx.Visible = False
    frmFYBX.dtgNx.Visible = True
    frmFYBX.cmdG.Visible = True
Case 54 '工程部
    frmFYBX.dtgBx.Columns("办公用品").Visible = True
    frmFYBX.dtgBx.Columns("通信费").Visible = True
    frmFYBX.dtgBx.Columns("市内交通费").Visible = True
    frmFYBX.dtgBx.Columns("市外交通费").Visible = True
    frmFYBX.dtgBx.Columns("住宿费").Visible = True
    frmFYBX.dtgBx.Columns("餐费").Visible = True
    'frmFYBX.dtgBx.Columns("工具").Visible = True
    frmFYBX.dtgBx.Columns("易耗").Visible = True
    frmFYBX.dtgBx.Columns("外劳").Visible = True
    frmFYBX.dtgBx.Columns("福利费").Visible = True
Case 70 '工程部
    frmFYBX.dtgBx.Columns("办公用品").Visible = True
    frmFYBX.dtgBx.Columns("通信费").Visible = True
    frmFYBX.dtgBx.Columns("市内交通费").Visible = True
    frmFYBX.dtgBx.Columns("市外交通费").Visible = True
    frmFYBX.dtgBx.Columns("住宿费").Visible = True
    frmFYBX.dtgBx.Columns("餐费").Visible = True
    'frmFYBX.dtgBx.Columns("工具").Visible = True
    frmFYBX.dtgBx.Columns("易耗").Visible = True
    frmFYBX.dtgBx.Columns("外劳").Visible = True
    frmFYBX.dtgBx.Columns("福利费").Visible = True
Case 55 '三金
    frmFYBX.dtgBx.Columns("三金").Visible = True
    frmFYBX.dtgBx.Columns("部门").Visible = True
    frmFYBX.dtgBx.Columns("区域").Visible = True
    frmFYBX.dtgBx.Columns("归属人").Visible = True
    frmFYBX.dtgBx.Columns("ywyuid").Width = 0
    frmFYBX.cmdG.Visible = True
        frmFYBX.dtgBx.Visible = True
    frmFYBX.dtgNx.Visible = False
Case 56 '公积金
    frmFYBX.dtgBx.Columns("公积金").Visible = True
    frmFYBX.dtgBx.Columns("部门").Visible = True
    frmFYBX.dtgBx.Columns("区域").Visible = True
    frmFYBX.dtgBx.Columns("归属人").Visible = True
    frmFYBX.dtgBx.Columns("ywyuid").Width = 0
    frmFYBX.cmdG.Visible = True
        frmFYBX.dtgBx.Visible = True
    frmFYBX.dtgNx.Visible = False
Case 58 '办事处公共费用
    frmFYBX.dtgBx.Columns("房租").Visible = True
    'frmFYBX.dtgBx.Columns("物业费").Visible = True
    frmFYBX.dtgBx.Columns("水电").Visible = True
    frmFYBX.dtgBx.Columns("电话").Visible = True
    frmFYBX.dtgBx.Columns("办公用品").Visible = True
    'frmFYBX.dtgBx.Columns("邮资").Visible = True
    frmFYBX.dtgBx.Columns("市场推广").Visible = True
    frmFYBX.dtgBx.Columns("人员招聘").Visible = True
    frmFYBX.dtgBx.Columns("快递费").Visible = True
    'frmFYBX.dtgBx.Columns("培训费").Visible = True
    frmFYBX.dtgBx.Columns("财务手续费").Visible = True
Case 59 '外来综合保险
    frmFYBX.dtgBx.Columns("综合保险").Visible = True
    frmFYBX.dtgBx.Columns("部门").Visible = True
    frmFYBX.dtgBx.Columns("区域").Visible = True
    frmFYBX.dtgBx.Columns("归属人").Visible = True
    frmFYBX.cmdG.Visible = True
    frmFYBX.dtgBx.Visible = True
    frmFYBX.dtgNx.Visible = False
Case 67 '房屋补贴
    frmFYBX.dtgBx.Columns("房屋补贴").Visible = True
    frmFYBX.dtgBx.Columns("部门").Visible = True
    frmFYBX.dtgBx.Columns("区域").Visible = True
    frmFYBX.dtgBx.Columns("归属人").Visible = True
Case 66 '生日
    frmFYBX.dtgBx.Columns("福利费").Visible = True
    frmFYBX.dtgBx.Columns("部门").Visible = True
    frmFYBX.dtgBx.Columns("区域").Visible = True
    frmFYBX.dtgBx.Columns("归属人").Visible = True
Case 72 '旅游费

    frmFYBX.dtgBx.Columns("部门").Visible = True
    frmFYBX.dtgBx.Columns("区域").Visible = True
    frmFYBX.dtgBx.Columns("归属人").Visible = True
    frmFYBX.dtgBx.Columns("归属人签字").Visible = True
    frmFYBX.dtgBx.Columns("签字时间").Visible = True
    frmFYBX.dtgBx.Columns("部门经理签字").Visible = True
    frmFYBX.dtgBx.Columns("签字日期").Visible = True
    frmFYBX.dtgBx.Columns("旅游费").Visible = True
    frmFYBX.dtgBx.Columns("日期").Visible = False
Case 84 '培训
    frmFYBX.dtgBx.Columns("培训费").Visible = True
    frmFYBX.dtgBx.Columns("部门").Visible = True
    frmFYBX.dtgBx.Columns("区域").Visible = True
    frmFYBX.dtgBx.Columns("归属人").Visible = True
    'frmFYBX.cmdG.Visible = True
        frmFYBX.dtgBx.Visible = True
    frmFYBX.dtgNx.Visible = False
Case 79 '新费用归属
frmFYBX.dtgBx.Columns("福利费").Visible = True
frmFYBX.dtgBx.Columns("房屋补贴").Visible = True
frmFYBX.dtgBx.Columns("旅游费").Visible = True
frmFYBX.dtgBx.Columns("高温费").Visible = True
frmFYBX.dtgBx.Columns("通信费").Visible = True
frmFYBX.dtgBx.Columns("市内交通费").Visible = True
frmFYBX.dtgBx.Columns("市外交通费").Visible = True
frmFYBX.dtgBx.Columns("运费").Visible = True
frmFYBX.dtgBx.Columns("住宿费").Visible = True
frmFYBX.dtgBx.Columns("部门团队费").Visible = True
frmFYBX.dtgBx.Columns("餐费").Visible = True
frmFYBX.dtgBx.Columns("招待费").Visible = True
frmFYBX.dtgBx.Columns("礼品费").Visible = True
frmFYBX.dtgBx.Columns("房租").Visible = True
frmFYBX.dtgBx.Columns("物业费").Visible = True
frmFYBX.dtgBx.Columns("水电").Visible = True
frmFYBX.dtgBx.Columns("电话").Visible = True
frmFYBX.dtgBx.Columns("办公用品").Visible = True
'frmFYBX.dtgBx.Columns("邮资").Visible = True
frmFYBX.dtgBx.Columns("市场推广").Visible = True
frmFYBX.dtgBx.Columns("人员招聘").Visible = True
frmFYBX.dtgBx.Columns("快递费").Visible = True
frmFYBX.dtgBx.Columns("培训费").Visible = True
frmFYBX.dtgBx.Columns("财务手续费").Visible = True
frmFYBX.dtgBx.Columns("团队建设费").Visible = True
frmFYBX.dtgBx.Columns("停车费").Visible = True
frmFYBX.dtgBx.Columns("车辆费").Visible = True
frmFYBX.dtgBx.Columns("公共停车费").Visible = True
frmFYBX.dtgBx.Columns("公共车辆费").Visible = True
'frmFYBX.dtgBx.Columns("工具").Visible = True
frmFYBX.dtgBx.Columns("易耗").Visible = True
frmFYBX.dtgBx.Columns("外劳").Visible = True
    frmFYBX.dtgBx.Columns("部门").Visible = False
    frmFYBX.dtgBx.Columns("区域").Visible = False
    frmFYBX.dtgBx.Columns("归属人").Visible = False
    frmFYBX.dtgBx.Columns("归属人签字").Visible = False
    frmFYBX.dtgBx.Columns("签字时间").Visible = False
    frmFYBX.dtgBx.Columns("部门经理签字").Visible = False
    frmFYBX.dtgBx.Columns("签字日期").Visible = False
frmFYBX.frmRen.Visible = True
Case 82 '内部结算
frmFYBX.dtgBx.Columns("快递费").Visible = True
frmFYBX.dtgBx.Columns("办公用品").Visible = True
frmFYBX.dtgBx.Columns("福利费").Visible = True
frmFYBX.dtgBx.Columns("停车费").Visible = True
frmFYBX.dtgBx.Columns("车辆费").Visible = True
    frmFYBX.dtgBx.Columns("部门").Visible = True
    frmFYBX.dtgBx.Columns("区域").Visible = True
    frmFYBX.dtgBx.Columns("归属人").Visible = True
'    frmFYBX.dtgBx.Columns("归属人签字").Visible = True
'    frmFYBX.dtgBx.Columns("签字时间").Visible = True
'    frmFYBX.dtgBx.Columns("部门经理签字").Visible = True
'    frmFYBX.dtgBx.Columns("签字日期").Visible = True
    frmFYBX.cmdG.Visible = True
   frmFYBX.dtgBx.Visible = True
   frmFYBX.dtgNx.Visible = False
End Select


frmFYBX.dtgBx.Refresh

End Sub

Public Sub fydBound(Bxid As String)
Dim tt As String
Dim oo As Integer
Dim Lcou As Integer
On Error Resume Next
Lcou = 0
frmFYBX.lblBh.Caption = Bxid
frmFYBX.cmdSave.Enabled = False
'记录打开日志
Call mod1.zhuDa(2, Bxid)

frmFYBX.Kd = False '非初次开单,以便保存时不生成员工签字日期

        tt = "fydOpen(" & Bxid & ")"
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
        frmFYBX.lblBh.Caption = mod1.HTP.Fields("BxId").Value
        frmFYBX.LblTrq.Caption = mod1.HTP.Fields("Trq").Value
        frmFYBX.comQy.Caption = mod1.HTP.Fields("qy").Value
        frmFYBX.lblBM.Caption = mod1.HTP.Fields("bm").Value
        frmFYBX.txtHg.Text = mod1.HTP.Fields("hG").Value
        frmFYBX.lblDx.Caption = mod1.HTP.Fields("hGD").Value
        frmFYBX.lblFR.Caption = mod1.HTP.Fields("fRQ").Value
        frmFYBX.lblLR.Caption = mod1.HTP.Fields("lRQ").Value
        frmFYBX.lblRq.Caption = mod1.HTP.Fields("QrQ").Value
        frmFYBX.txtQc.Text = mod1.HTP.Fields("QMin").Value
        frmFYBX.lblNlb.Caption = mod1.HTP.Fields("Nlb").Value
        frmFYBX.txtBz.Text = mod1.HTP.Fields("Bz").Value
        frmFYBX.lblBt.Caption = mod1.HTP.Fields("Fbt").Value

        frmFYBX.txtCwBZ.Text = mod1.HTP.Fields("CWBZ").Value
        frmFYBX.lblLc.Caption = mod1.HTP.Fields("LC").Value
        frmFYBX.lblLcRen.Caption = mod1.HTP.Fields("LCren").Value
        frmFYBX.lblLcUid.Caption = mod1.HTP.Fields("LCuid").Value
        frmFYBX.lblYwy.Caption = mod1.HTP.Fields("ywy").Value  '单子所属人
        frmFYBX.lblUid.Caption = mod1.HTP.Fields("Uid").Value
        frmFYBX.lblFwid.Caption = mod1.HTP.Fields("Fwid").Value '当前对应NewFuwu表的ID
        Lcou = mod1.HTP.Fields("Lcou").Value '流程总数
        frmFYBX.lblYqf.Caption = mod1.HTP.Fields("yqf").Value  '业务审核的各人员是否都签字
        frmFYBX.lblGui.Caption = mod1.HTP.Fields("GRen").Value '归属人
        frmFYBX.lblGuid.Caption = mod1.HTP.Fields("Grid").Value
        frmFYBX.optFp1.Value = mod1.HTP.Fields("fp").Value
        frmFYBX.lblNewF.Caption = mod1.HTP.Fields("newF").Value
        If mod1.HTP.Fields("czf").Value = True Then '是否显示附加签名
            frmFYBX.frmZQ.Visible = True
            frmFYBX.cmdFQ.Caption = mod1.HTP.Fields("zjin").Value
            frmFYBX.lblFT.Caption = mod1.HTP.Fields("tc").Value
        Else
            frmFYBX.frmZQ.Visible = False
        End If
        If frmFYBX.optFp1.Value = False Then
            frmFYBX.optFp2.Value = True
            frmFYBX.txtFP.Text = mod1.HTP.Fields("fpnr").Value
        End If
        '按老版或新版显示不同的签字按钮
        If IsNull(mod1.HTP.Fields("Lcou").Value) = True Then
            frmFYBX.frmQm.Visible = True
            frmFYBX.cmdBxr.Caption = mod1.HTP.Fields("yWy").Value
            frmFYBX.cmdJc.Caption = mod1.HTP.Fields("Jian").Value
            frmFYBX.cmdJl.Caption = mod1.HTP.Fields("JinLi").Value
            frmFYBX.cmdZj.Caption = mod1.HTP.Fields("zJin").Value
            frmFYBX.lblTa.Caption = mod1.HTP.Fields("ta").Value
            frmFYBX.lblTb.Caption = mod1.HTP.Fields("tb").Value
            frmFYBX.lblTC.Caption = mod1.HTP.Fields("tc").Value
            frmFYBX.lblTd.Caption = mod1.HTP.Fields("td").Value
        Else                                           '新版
            frmFYBX.frmNewQ.Visible = True
            'Call ModBx.AddLcBut(mod1.HTP.Fields("Nlb").Value)

            tt = "FydQmOpen('" & frmFYBX.lblBh.Caption & "'," & 23 & ")" '23为workBl中的报销单事务编号
            mod1.HTT.Close
            mod1.HTT.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
            mod1.HTT.MoveFirst
            For oo = 0 To mod1.HTT.RecordCount - 1
                frmFYBX.cmdQm(oo).Caption = mod1.HTT.Fields("QRen").Value
                frmFYBX.lblTm(oo).Caption = mod1.HTT.Fields("QRQ").Value
                mod1.HTT.MoveNext
            Next
        End If
        If frmFYBX.txtQc.Text <> "" Then
            frmFYBX.txtQc.PasswordChar = ""
            frmFYBX.txtQc.Enabled = False
        Else
            frmFYBX.txtQc.PasswordChar = "*"
            frmFYBX.txtQc.Enabled = True
        End If
        If Val(frmFYBX.lblBh.Caption) > 124571 Then
            frmFYBX.frmG.Visible = True
        End If

        
        '打开费用总表
    tt = "FydMxOpen(" & Val(Bxid) & ")"
 
    Call ModBx.dtgKj(frmFYBX.lblNlb.Caption)
    If IsNull(mod1.HTP.Fields("lcou").Value) = True And frmFYBX.lblNlb.Caption = 9 Then '老版中房屋补贴
        frmFYBX.dtgBx.Columns("房屋补贴").Visible = True
        frmFYBX.dtgBx.Columns("部门").Visible = True
        frmFYBX.dtgBx.Columns("区域").Visible = True
        frmFYBX.dtgBx.Columns("归属人").Visible = True
    End If
        frmFYBX.cmdAdd.Visible = False
        frmFYBX.cmdDel.Visible = False
        frmFYBX.cmdSave.Enabled = False
        frmFYBX.dtgBx.AllowUpdate = False
        
       
    If mod1.HTP.Fields("lc") = 1 Or mod1.HTP.Fields("lc") = 0 Then   '如果是初次开单
        frmFYBX.cmdMod.Enabled = True
        frmFYBX.adoF2.Recordset.Close
        frmFYBX.adoF2.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdStoredProc
    Else
        frmFYBX.cmdMod.Enabled = False
        frmFYBX.adoF2.Recordset.Close
        frmFYBX.adoF2.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
    End If

        Set frmFYBX.dtgBx.DataSource = frmFYBX.adoF2
        tt = "Select atime as 日期,khmc as 报销内容,sj as 三金,fwbt as 房屋补贴,lyf as 旅游费,gwf as 高温费,txf as 通信费,njtf as 市内交通费,wjtf as 市外交通费," & _
        "tcf as 停车费,clf as 车辆费,yf as 运费,zcf as 住宿费,bmtd as 部门团队费,cf as 餐费,ZDF as 招待费,LPF as 礼品费,fz as 房租,WYF as 物业费," & _
        "sd as 水电,DW as 电话,BGYP as 办公用品,YZ as 邮资,SZTG as 市场推广,RYZP as 人员招聘,KDF as 快递费,PXF as 培训费,CWSX as 财务手续费,TDJS as 团队建设费," & _
        "GTCF as 公共停车费,GCLF as 公共车辆费,gg as 工具,yH as 易耗,wl as 外劳,qtf as 福利费,gjj as 公积金,zhbx as 综合保险,jtbt as 交通补贴,zwbt as 驻外津贴,gwbt as 岗位补贴,bm as 部门,qy as 区域,ywy as 姓名," & _
        "bid,gzdh as 出租车注明,xg as 合计,qrq as 签收日期,GongF,GBM from fyBx where Bxid=" & Val(Bxid) & " order by bm,bid"
        frmFYBX.Fmx.Close
        frmFYBX.Fmx.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
        Call ModBx.DiZ










        
    If (mod1.DName = "文静" Or mod1.DName = "乔继敏") And frmFYBX.lblLc.Caption > 1 Then
        frmFYBX.txtCwBZ.Enabled = True
        frmFYBX.txtCwBZ.Locked = False
        frmFYBX.cmdSave.Enabled = True
        frmFYBX.txtBz.Locked = True
    End If
    
        '打开流程按钮.
        Call OpenAN
    'If mod1.Bq2 = True And frmFYBX.txtQM = "" And frmFYBX.lblLcRen.Caption = mod1.DName Then
    'If frmFYBX.lblLc.Caption = Lcou Then '如果到了流程最后,则可以密码签收
    If mod1.Bq2 = True Then
        frmFYBX.txtQc.Enabled = True
    Else
        frmFYBX.txtQc.Enabled = False
    End If
    If Val(frmFYBX.lblNlb.Caption) = 79 Then
        frmFYBX.cmdMod.Enabled = True
    End If
    frmFYBX.frmEd.Visible = False
    frmFYBX.cmdG.Visible = False
    Call frmFYBX.QMBound(Val(Bxid))
End Sub
















Public Sub FyQing() '营销部报销单清空
Dim oo As Integer
On Error Resume Next
    frmFYBX.frmNewQ.Visible = False
    frmFYBX.lblBh.Caption = ""
    frmFYBX.comQy.Caption = "上海"
    frmFYBX.txtHg.Text = ""
    frmFYBX.lblDx.Caption = ""
    frmFYBX.lblFR.Caption = ""
    frmFYBX.lblLR.Caption = ""
    frmFYBX.lblRq.Caption = ""
    frmFYBX.cmdBxr.Caption = ""
    frmFYBX.cmdJc.Caption = ""
    frmFYBX.cmdJl.Caption = ""
    frmFYBX.cmdZj.Caption = ""
    frmFYBX.comDQ.Text = ""
    frmFYBX.txtQc.Text = ""
    frmFYBX.txtCwBZ.Text = ""
    frmFYBX.txtBz.Text = ""
    frmFYBX.lblTa.Caption = ""
    frmFYBX.lblTb.Caption = ""
    frmFYBX.lblTC.Caption = ""
    frmFYBX.lblTd.Caption = ""
    frmFYBX.LblTrq.Caption = ""
    frmFYBX.lblNlb.Caption = ""
    frmFYBX.frmQm.Visible = False
    frmFYBX.frmNewQ.Visible = False
    frmFYBX.frmYf.Visible = False
    frmFYBX.frmWd.Visible = False
    frmFYBX.lblLc.Caption = ""
    frmFYBX.lblLcRen.Caption = ""
    frmFYBX.lblLcUid.Caption = ""
    frmFYBX.lblBt.Caption = ""
    For oo = 5 To 0 Step -1
        Unload frmFYBX.lblQM(oo)
        Unload frmFYBX.cmdQm(oo)
        Unload frmFYBX.lblTm(oo)
    Next
    frmFYBX.lblQM(0).Caption = "报销人"
    frmFYBX.cmdQm(0).Caption = ""
    frmFYBX.lblTm(0).Caption = ""
    frmFYBX.txtCwBZ.Enabled = False '财务备注只能在财务审核时能编辑
    frmFYBX.lblYwy.Caption = "" '单子所属人
    frmFYBX.lblUid.Caption = ""
    frmFYBX.lblFwid.Caption = "" '当前对应NewFuwu表的ID
    frmFYBX.lblYqf.Caption = "" '业务审核的各人员是否都签字
    frmFYBX.frmRen.Visible = False
    frmFYBX.lblGui.Caption = ""
    frmFYBX.lblGuid.Caption = ""
    frmFYBX.cmdGui.Visible = False
    frmFYBX.cmdDao.Visible = False
    frmFYBX.optFp1.Value = False
    frmFYBX.optFp2.Value = False
    frmFYBX.txtFP.Text = ""
    frmFYBX.lblBid.Caption = ""
    frmFYBX.lblNewF.Caption = ""
    frmFYBX.lblTx.Visible = False
    frmFYBX.lblGZDH.Visible = False
    frmFYBX.txtGZDH.Visible = False
    frmFYBX.frmZQ.Visible = False
    frmFYBX.cmdFQ.Caption = ""
    frmFYBX.lblFT.Caption = ""
    frmFYBX.lbl1.Caption = "" '公共费用
    frmFYBX.lbl2.Caption = "" '个人费用
    frmFYBX.frmG.Visible = False
    frmFYBX.txtBm.Text = ""
    Call frmFYBX.dtgPFF
End Sub
Public Sub AddLcBut(Nlb As Integer)  '添加流程签字按钮
Dim tt As String
Dim oo As Integer
On Error Resume Next
    tt = "lcBut(" & Nlb & ")"
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
    mod1.HTP.MoveFirst
    mod1.HTP.MoveNext '第一个数组按钮不用添加,所以,跳到下一记录
    For oo = 1 To mod1.HTP.RecordCount - 1
        Load frmFYBX.lblQM(oo)
        Load frmFYBX.cmdQm(oo)
        Load frmFYBX.lblTm(oo)
        frmFYBX.lblQM(oo).Caption = mod1.HTP.Fields("LNR").Value
        frmFYBX.lblQM(oo).Visible = True
        frmFYBX.lblQM(oo).Left = frmFYBX.lblQM(oo - 1).Left + 1100
        frmFYBX.cmdQm(oo).Caption = ""
        frmFYBX.lblTm(oo).Caption = ""
        frmFYBX.cmdQm(oo).Visible = True
        frmFYBX.lblTm(oo).Visible = True
        frmFYBX.cmdQm(oo).Left = frmFYBX.cmdQm(oo - 1).Left + 1100
        frmFYBX.lblTm(oo).Left = frmFYBX.lblTm(oo - 1).Left + 1100
        mod1.HTP.MoveNext
    Next

'添加进QMRZ表
'tt = "QMrzOpen('豪曼')"
'mod1.HTT.Close
'mod1.HTT.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdStoredProc
'mod1.HTP.MoveFirst
'Do While Not mod1.HTP.EOF
'    mod1.HTT.AddNew "Qlabel", mod1.HTP.Fields("LNR").Value
'    mod1.HTT.Update "BTZ", 23  '报销单
'    mod1.HTT.Update "QDBh", frmFYBX.lblBh.Caption '编号
'    mod1.HTT.Update "Zid", mod1.HTP.Fields("zid").Value '顺序
'    If mod1.HTP.Fields("mid").Value = 38 Or mod1.HTP.Fields("mid").Value = 43 Or _
'       mod1.HTP.Fields("mid").Value = 48 Then                                    '是否为业务审核明细签字
'       mod1.HTT.Update "MXQF", 1
'    End If
'    mod1.HTT.UpdateBatch
'    mod1.HTP.MoveNext
'Loop
        Set mod1.cmd = CreateObject("adodb.command")
        mod1.cmd.ActiveConnection = mod1.cc
        mod1.cmd.CommandText = "QMRZAdd"
        mod1.cmd.CommandType = adCmdStoredProc
        mod1.cmd.Parameters("@NLb") = Nlb
        mod1.cmd.Parameters("@btz") = mod1.BTZ
        mod1.cmd.Parameters("@QDBH") = frmFYBX.lblBh.Caption '编号
        mod1.cmd.Execute
        Set cmd = Nothing
'        If Nlb = 79 Then
'            frmFYBX.lblQM(0).Caption = "归属人"
'        End If
End Sub

Public Sub OpenAN()
Dim tt As String
Dim oo As Integer
On Error Resume Next
    For oo = 10 To 1 Step -1
        Unload frmFYBX.cmdQm(oo)
        Unload frmFYBX.lblQM(oo)
        Unload frmFYBX.lblTm(oo)
    Next

      'tt = "qmrzOpen(" & mod1.BTZ & ",'" & frmFYBX.lblBh.Caption & "')"
      tt = "qmrzOpen(23,'" & frmFYBX.lblBh.Caption & "')"
      Set mod1.HTP = CreateObject("adodb.recordset")
      mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
      mod1.HTP.MoveFirst
      frmFYBX.lblQM(0).Caption = mod1.HTP.Fields("QLabel").Value
        If mod1.HTP.Fields("xf").Value = True Then
            frmFYBX.cmdQm(0).Caption = mod1.HTP.Fields("Qren").Value
            frmFYBX.lblTm(0).Caption = mod1.HTP.Fields("QRQ").Value
        Else
            frmFYBX.cmdQm(0).Caption = ""
            frmFYBX.lblTm(0).Caption = ""
        End If
      frmFYBX.cmdQm(0).Tag = mod1.HTP.Fields("zid").Value
      mod1.HTP.MoveNext
      For oo = 1 To mod1.HTP.RecordCount - 1
        Load frmFYBX.lblQM(oo)
        frmFYBX.lblQM(oo).Caption = ""
        Load frmFYBX.cmdQm(oo)
        frmFYBX.cmdQm(oo).Caption = ""
        Load frmFYBX.lblTm(oo)
        frmFYBX.lblTm(oo).Caption = ""
        frmFYBX.lblQM(oo).Caption = mod1.HTP.Fields("QLabel").Value
        If mod1.HTP.Fields("xf").Value = True Then
            frmFYBX.cmdQm(oo).Caption = mod1.HTP.Fields("Qren").Value
            If frmFYBX.cmdQm(oo).Caption = "南京办经理" Then
                frmFYBX.cmdQm(oo).Caption = "南京办经理"
            End If
            frmFYBX.lblTm(oo).Caption = mod1.HTP.Fields("QRQ").Value
        End If

        frmFYBX.cmdQm(oo).Tag = mod1.HTP.Fields("zid").Value
        frmFYBX.lblQM(oo).Visible = True
        frmFYBX.cmdQm(oo).Visible = True
        frmFYBX.lblTm(oo).Visible = True
        frmFYBX.lblQM(oo).Left = frmFYBX.lblQM(oo - 1).Left + 1100
        frmFYBX.cmdQm(oo).Left = frmFYBX.cmdQm(oo - 1).Left + 1100
        frmFYBX.lblTm(oo).Left = frmFYBX.lblTm(oo - 1).Left + 1100
        mod1.HTP.MoveNext
        
     Next
End Sub

Public Sub DiZ()
Dim oo As Integer
Dim rr As Integer
Dim F1 As Single '公共费用合计
Dim F2 As Single '个人费用合计
On Error Resume Next
F1 = 0: F2 = 0
        frmFYBX.Fmx.Requery
        'Set frmFYBX.dtgNx.DataSource = frmFYBX.Fmx
        
If frmFYBX.Fmx.RecordCount = 0 Then
    Set frmFYBX.dtgNx.DataSource = frmFYBX.Fmx
    frmFYBX.dtgNx.Rows = 2
    frmFYBX.dtgNx.FixedRows = 0
    frmFYBX.dtgNx.FixedRows = 1

Else
    frmFYBX.dtgNx.Rows = 2
    frmFYBX.dtgNx.FixedRows = 1
    Set frmFYBX.dtgNx.DataSource = frmFYBX.Fmx
End If
        '显示有值字段
        For oo = 3 To 40
            frmFYBX.dtgNx.ColWidth(oo) = 0
        Next


        For oo = 3 To 40
            rr = 1
            frmFYBX.dtgNx.Col = oo
            Do While Not rr >= frmFYBX.dtgNx.Rows
                frmFYBX.dtgNx.Row = rr
                If Val(frmFYBX.dtgNx.Text) > 0 Then
                    frmFYBX.dtgNx.ColWidth(oo) = 1000
                    
                    frmFYBX.dtgNx.Col = 48
                    If Val(frmFYBX.dtgNx.Text) = 1 Then
                        frmFYBX.dtgNx.Col = oo
                        frmFYBX.dtgNx.CellForeColor = &HFF&
                        F1 = F1 + Val(frmFYBX.dtgNx.Text)
                    ElseIf Val(frmFYBX.dtgNx.Text) = 2 Then
                        frmFYBX.dtgNx.Col = oo
                        frmFYBX.dtgNx.CellForeColor = &HC00000
                        F2 = F2 + Val(frmFYBX.dtgNx.Text)
                    End If
                    
                    'Exit Do
                End If
                rr = rr + 1
            Loop
        Next
        frmFYBX.lbl1.Caption = F2: frmFYBX.lbl2.Caption = F1
'        If frmFYBX.dtgNx.ColWidth(3) = 1005 Or frmFYBX.dtgNx.ColWidth(36) = 1005 Then
'            frmFYBX.dtgNx.ColWidth(36) = 1000
'            frmFYBX.dtgNx.ColWidth(3) = 0
'        End If
        frmFYBX.dtgNx.FixedRows = 0
        frmFYBX.dtgNx.MergeCol(1) = True
        frmFYBX.dtgNx.MergeCol(2) = True
        frmFYBX.dtgNx.MergeCol(41) = True
        frmFYBX.dtgNx.MergeCol(42) = True
        frmFYBX.dtgNx.MergeCol(43) = True
        frmFYBX.dtgNx.MergeCells = 3
        frmFYBX.dtgNx.FixedRows = 1
        'If frmFYBX.lblBm.Caption = "工程部" Then
            frmFYBX.dtgNx.ColWidth(45) = 1000
        'Else
            'frmFYBX.dtgNx.ColWidth(41) = 0
        'End If
        If frmFYBX.lblNlb.Caption = 35 Then
            frmFYBX.dtgNx.ColWidth(45) = 0
            'frmFYBX.dtgNx.ColWidth(40) = 1000
        Else
            frmFYBX.dtgNx.ColWidth(40) = 0
        End If
        
        
End Sub
