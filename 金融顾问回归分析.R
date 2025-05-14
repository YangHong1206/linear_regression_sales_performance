  #### Libraries ####
  library(lubridate)
  library(olsrr)
  library(tidyverse)
  library(readxl)
  library(agridat)
  library(writexl)
  library(broom) #linear regression model summary
  library(glue)
  library(car)
  library(tidyplots)
  library(corrplot)
  library(openxlsx)
  library(reshape2)
  library(rmarkdown)
  library(caret) #逐步回归-找最优的变量
  library(gptstudio) #使用chatgpt
  library(ggthemes)#使用tableau的颜色
  library(esquisse) #use,esquisser 拖放绘图
  library(ggThemeAssist) # 交互式主题编辑
  library(colourpicker)
  library(tsviz)#时间序列可视化
  library(cols4all)#颜色版 c4a_gui()交互界面

  
  #### setting path ####
  setwd("D:/Chrome_download")
 
  
  
  #### data—Sources are manually renewable ####
  
  rawdata <- read.csv("历史订单明细脱敏版.csv")
  pre_apr <- read.csv("预审订单明细.csv")
  act_channel<- read_xlsx("渠道管理列表.xlsx")
  all_channel <- read_xlsx("渠道管理列表.xlsx")
  channel <- read_excel("渠道管理列表.xlsx")
  visit_record <- read_xlsx("店面拜访.xlsx")
  华东员工<- read_xlsx("2025-05-04华东区花名册.xlsx")
  历史累积拜访 <- read.xlsx("累积拜访数据源2112_2203.xlsx")
  目标额台<- read_xlsx("D:/personal doc/杨宏/分公司助理/日报/目标转化/福州消费融目标.xlsx")
  #礼品投放<- read.xlsx("市场礼品申请及复盘.xlsx", sheet = "福州场礼品投放渠道", startRow = 2 )
  #vist2403 <- read.xlsx("店面拜访.xlsx")

  

  #已经有新的累积拜访历史数据，新的一个月需要先将拜访表修改开始时间格式为日期,
  #拜访用时间需要修改为int格式
  
  #月拜访记录 <- read_xlsx("店面拜访.xlsx")
  
  #本月拜访记录 <- 月拜访记录[1:12]
  
  #累积拜访数据源2112_2203 <- rbind( 
    #历史累积拜访, 拜访记录)
  
  #write.xlsx(累积拜访数据源2112_2203,"D:/Chrome_download/累积拜访数据源2112_2203.xlsx")
  
  
  ####Date variants ####
  
  现在年月 <- paste( year(today()),month(today()),sep = "")
  
  #未成交客户跟进时间
  #未成交客户跟进开始日期 <- paste( year(today()),month(today()),day(1),sep = "/")
  

  write_excel_csv(visit_record, "visit_record.csv")  # 添加UTF-8 BOM
  write_excel_csv(act_channel, "channel.csv")
  write_excel_csv(rawdata, "rawdata.csv")  # 添加UTF-8 BOM
  write_excel_csv(pre_apr, "pre_apr.csv")
  
  
  # readr::write_csv(visit_record,"visit_record.csv")
  # readr::write_csv(act_channel, "channel.csv")
  
  #Background_of_ggplot----
  
  
  # Functions----
  
  
  #get rid of ()
  fn_drop_parentheses <- function(clean_parenses_a)
  {
    clean_parenses_a <- gsub("\\(.*\\)", "", clean_parenses_a)
    clean_parenses_a <- gsub("\\（.*\\）", "", clean_parenses_a)
    return(clean_parenses_a)
  }
  
  #adjust ymd
  fn_format_ymd <- function(date_a)
  {
    date_a <- as.Date(date_a, format = "%Y-%m-%d")
    return(date_a)
  }
  
  #adjust hours
  fn_format_h <- function(time_a)
  {
    time_a <- format(as.POSIXct(time_a, format="%Y-%m-%d %H:%M:%S"),"%H")
    return(time_a)
  }
  

  
  #### Tidy data ####
  
  rawdata <- rawdata %>% 
    mutate(
      融资金额 = 融资金额 / 10000,
      进件日期 = ymd(进件日期),
      合同生效日期 = ymd(合同生效日期),
      金融经理名称 = fn_drop_parentheses(金融经理名称),
      金融顾问名称 = fn_drop_parentheses(金融顾问名称),
      #pad.by.timedd = pad_by_time(合同生效日期, )
    ) %>% 
    arrange(desc(合同生效日期))
  

  
  pre_apr <- pre_apr %>%
    mutate(
      进件时点 = fn_format_h(进件时间),
      预审提交时间 = fn_format_ymd(预审提交时间),
      金融经理 = fn_drop_parentheses(金融经理),
      金融顾问 = fn_drop_parentheses(金融顾问),
      进件时间 = fn_format_ymd(进件时间)
      
    )

  
  
  
  act_channel <- act_channel %>%
    rename_at('负责员工(金融顾问)', ~'负责员工') %>%    
    mutate(负责员工 = fn_drop_parentheses(负责员工),
           #负责员工 = gsub("\\（.*\\）","",负责员工),
           业务上级 = fn_drop_parentheses(业务上级)
           #业务上级 = gsub("\\（.*\\）","",业务上级)
           )
  
  channel <- channel %>%
    rename_at('负责员工(金融顾问)', ~'负责员工') %>% 
    mutate(负责员工 = fn_drop_parentheses(负责员工),
           业务上级 = fn_drop_parentheses(业务上级),
           初始生效日期 = ymd(初始生效日期),
           渠道关停日期 = ymd(渠道关停日期)
    )

  all_channel <- all_channel %>%
    rename_at('负责员工(金融顾问)', ~'负责员工') %>% 
    mutate(负责员工 = fn_drop_parentheses(负责员工),
           业务上级 = fn_drop_parentheses(业务上级),
           初始生效日期  = fn_format_ymd(初始生效日期),
           渠道关停日期 = fn_format_ymd(渠道关停日期)
    )
  
  temp_all_channel <- all_channel %>% 
    mutate(初始生效年月 = format(初始生效日期, "%Y-%m"),
           渠道关停年月 = format(渠道关停日期, "%Y-%m"))
  

  visit_record <- visit_record %>%
    mutate(
      拜访开始时间 = as.POSIXlt(拜访开始时间),
      拜访结束时间 = as.POSIXlt(拜访结束时间),
      拜访用时 = as.numeric(拜访用时),
      拜访人 = fn_drop_parentheses(拜访人),
      业务上级 = fn_drop_parentheses(业务上级)
    )
  
  
  历史累积拜访 <- 历史累积拜访 %>% 
    mutate( 
      拜访结束年月日 = date( as.Date(拜访结束时间) ), 
      拜访结束年月 = as.character(year(as.Date(拜访结束时间))*100 +
                     month(as.Date(拜访结束时间)))
      )

  #渠道经理顾问表----
  渠道经理顾问表 <- rawdata %>%
    filter(合同生效日期 >= today() - 90) %>%
    arrange(desc(合同生效日期)) %>%
    select(渠道代码, 店面主体, 金融经理名称, 金融顾问名称) %>%
    unique()
  
  #消费融所有进件渠道生效记录----
  
  
  
#金融顾问回归分析----
#数据准备
  if(T){

  #合并历史拜访数据
    file_paths <- c("vist2401.xlsx","店面拜访 (vist2402).xlsx","店面拜访 (vist2403).xlsx",
                    "店面拜访 (vist2404).xlsx","店面拜访 (vist2405).xlsx","店面拜访 (vist2406).xlsx",
                    "店面拜访 (vist2407).xlsx","店面拜访 (vist2408).xlsx","店面拜访 (vist2409).xlsx",
                    "店面拜访 (vist2410).xlsx","店面拜访 (vist2411).xlsx","店面拜访 (vist2412).xlsx",
                    "店面拜访 (vist2501).xlsx","店面拜访 (vist2502).xlsx","店面拜访 (vist2503).xlsx",
                    "店面拜访 (vist2504).xlsx","店面拜访 (vist2301).xlsx","店面拜访 (vist2302).xlsx"
                    ,"店面拜访 (vist2303).xlsx","店面拜访 (vist2304).xlsx","店面拜访 (vist2305).xlsx"
                    ,"店面拜访 (vist2306).xlsx","店面拜访 (vist2307).xlsx","店面拜访 (vist2308).xlsx"
                    ,"店面拜访 (vist2309).xlsx","店面拜访 (vist2310).xlsx","店面拜访 (vist2311).xlsx"
                    ,"店面拜访 (vist2312).xlsx")
    
    # 循环读取并合并
    新版本拜访数据2301_2504 <- lapply(file_paths, read_xlsx) %>% 
      bind_rows(.id = "来源文件") %>%  # 可选：用.id参数标记来源
      select(-来源文件)
    
  

  

  vist2112_2210 <- read.xlsx("店面拜访记录202112-202210.xlsx",colNames = TRUE)
  
  员工职位<- 华东员工 %>% 
    select(姓名,职位名称)
  str(vist2112_2210)
  
  old_vist_history <- vist2112_2210 %>% 
    left_join(员工职位, join_by("拜访人"=="姓名")) %>% 
    mutate(拜访开始时间 = paste(开始日期,开始时间),
           拜访结束时间 = paste(结束日期,结束时间)) %>% 
    select(所属大区, 所属分公司, 业务上级, 拜访人, 职位名称, 店面编号, 
           店面名称,主体名称, 拜访开始时间, 拜访结束时间, 拜访是否达标,
           '拜访时长（分钟）') %>% 
    rename(大区 = 所属大区, 销售三级类 = 职位名称,
           是否达标 = 拜访是否达标, 拜访用时 = '拜访时长（分钟）')
  
  
  #累积拜访数据源<- rbind(old_vist_history,新版本拜访数据2301_2504)
  
  #累积拜访数据源<- 累积拜访数据源 %>%
  累积拜访数据源 <- 新版本拜访数据2301_2504 %>%
    mutate(
      业务上级  = gsub("\\(.*\\)", "", 业务上级) ,
      业务上级  = gsub("\\（.*\\）", "", 业务上级),
      拜访人  = gsub("\\(.*\\)", "", 拜访人)  ,
      拜访人  = gsub("\\（.*\\）", "", 拜访人),
      拜访用时 = as.integer(拜访用时),
      拜访开始时间 = as.Date(拜访开始时间, format = "%Y-%m-%d %H:%M:%S"),
      拜访开始时间 = as.Date(拜访开始时间, format = "%Y-%m-%d %H:%M:%S")
    ) %>% 
    left_join(
      华东员工 %>% select(`员工 ID`, 姓名),  # 右表数据
      by = c("拜访人" = "姓名")  # 正确指定列对应关系
    )
  
    
  write.xlsx(累积拜访数据源,"D:/Chrome_download/累积拜访数据源2301-2504.xlsx")
  #导出历史拜访数据数据源
  
  
  
  顾问每月拜访数和时间 <- 累积拜访数据源 %>%
    mutate(年月 = format(拜访开始时间,"%Y%m")) %>%
    rename(员工编号 = `员工 ID`) %>% 
    group_by(年月, 员工编号, 拜访人) %>% 
    summarize(拜访门店数 = length(主体名称),
              累计拜访用时 = sum(拜访用时),
              拜访中位数med = median(拜访用时))
  }


  
    
#消费融所有顾问预审进件生效记录----
if(T){

顾问预审 <- pre_apr %>%
  arrange(desc(预审提交时间)) %>%
  mutate(预审月 = format(预审提交时间, "%Y%m")) %>%
  group_by(金融顾问, 预审月, 车辆类型, 渠道二级科目名称, 店面城市) %>%
  dplyr::summarise(预审量 = length(申请编号))



view <- rawdata %>%
  mutate(
    进件月 = format(进件日期, "%Y%m"),
    生效月 = format(合同生效日期, "%Y%m"),
    进件订单 = case_when(!is.na(进件日期) ~ 1, .default = NULL),
    生效单量 = case_when(!is.na(合同生效日期) & 申请状态 != "订单已取消"
                     ~ 1, .default = NULL),
    生效额 = case_when(!is.na(合同生效日期) & 申请状态 != "订单已取消"
                    ~ 融资金额, .default = NULL)
  )


#write.xlsx(view, "D:/Chrome_download/view.xlsx")


顾问全进件 <- view %>%
  group_by(金融顾问名称, 进件月, 车辆类型, 渠道二级科目, 店面城市) %>%
  dplyr::summarise(进件量 = sum(进件订单))

顾问全生效额 <- view %>%
  group_by(金融顾问名称, 生效月, 车辆类型, 渠道二级科目, 店面城市) %>%
  dplyr::summarise(生效金额 = sum(生效额, na.rm = TRUE)) %>%
  na.omit()


全顾问活跃门店数 <- view %>%
  group_by(金融顾问名称, 生效月, 车辆类型, 渠道二级科目, 店面城市) %>%
  summarise(活跃渠道数 = length(unique(店面主体))) %>%
  na.omit()

#write.xlsx(全顾问活跃门店数,"D:/Chrome_download/全顾问活跃门店数.xlsx")


全顾问生效量 <- view %>%
  group_by(金融顾问名称, 生效月, 车辆类型, 渠道二级科目, 店面城市) %>%
  dplyr::summarise(生效量 = length(生效单量)) %>%
  na.omit() %>%
  left_join(顾问全进件, join_by("金融顾问名称" == "金融顾问名称", "生效月" == "进件月",
                           "车辆类型" == "车辆类型", "渠道二级科目" == "渠道二级科目", "店面城市" == "店面城市")) %>%
  left_join(顾问预审, join_by("金融顾问名称" == "金融顾问", "生效月" == "预审月",
                          "车辆类型" == "车辆类型", "渠道二级科目" == "渠道二级科目名称", "店面城市" == "店面城市")) %>%
  left_join(顾问全生效额, join_by("金融顾问名称" == "金融顾问名称", "生效月" == "生效月",
                            "车辆类型" == "车辆类型", "渠道二级科目" == "渠道二级科目", "店面城市" == "店面城市")) %>%
  left_join(全顾问活跃门店数, join_by("金融顾问名称" == "金融顾问名称", "生效月" == "生效月",
                              "车辆类型" == "车辆类型", "渠道二级科目" == "渠道二级科目", "店面城市" == "店面城市")) %>% 

   left_join(顾问每月拜访数和时间, join_by("金融顾问名称" == "拜访人", "生效月" == "年月")) %>% 
  mutate_if(., is.numeric, ~ replace(., is.na(.), 0))


write.xlsx(全顾问生效量, "D:/Chrome_download/全顾问生效量.xlsx")

remove(顾问预审, 顾问全进件, 顾问全生效额, 全顾问活跃门店数)



#变量关系----
顾问销售表现数据_numeric <- 全顾问生效量 %>% 
  filter(金融顾问名称 != "陈伟彬") %>%
  ungroup() %>%  # 关键：解除分组状态
  select(-金融顾问名称, -员工编号)  # 现在可以正常删除列


summary(顾问销售表现数据_numeric)#检查缺失值

cor_matrix <- cor(顾问销售表现数据_numeric[,-1:-4], use = "complete.obs")  # 处理缺失值
range(cor_matrix, na.rm = TRUE)  

corrplot(cor_matrix, 
         type = "upper",
         method = "color",
         tl.col = "black",
         addCoef.col = "black",  # 显示相关系数
         number.cex = 0.8)

}

  
  
#多元回归，multiple linear regreesion----
  

  销售顾问lrmod <- lm(生效金额 ~ 
                    预审量 + 进件量 + 活跃渠道数 + 拜访门店数 , 顾问销售表现数据_numeric,   ) #剔除 +累计拜访用时 +拜访中位数med
  
  summary(销售顾问lrmod)
  
  
  
#诊断残差 estimating residual error----
  plot(销售顾问lrmod$residuals)
  
  mean(销售顾问lrmod$residuals)
  # 2.079371e-15 非常接近0 满足零均值准则
  
  #残差正态性 #residual error distribution
  
  
  ols_plot_resid_hist(销售顾问lrmod)
  
  
  #残差和拟合值的关系图，检验异方差性
  ols_plot_resid_fit(销售顾问lrmod)
  
  
  #残差自相关
  
  #分析有影响的点

  durbinWatsonTest(销售顾问lrmod)
  

  #得到DW STATISTIC 检验统计量为1.28854， P value 是0
  #模型残差表现出正相关
  
  #分析有影响的点
  
  # 包含Cook's距离、帽子矩阵值、学生化残差
  influenceIndexPlot(
    销售顾问lrmod,
    main = "影响点诊断",
    id = list(n = 10)  # 自动标注前5个高影响点
  )
  
  ols_plot_cooksd_chart(销售顾问lrmod)
  
  
  
  销售顾问lrmod_cooks_outliers <- ols_plot_cooksd_chart(销售顾问lrmod)$outliers
  
  if(!is.null(销售顾问lrmod_cooks_outliers)) {
    # 降序排列
    arranged_outliers <- 销售顾问lrmod_cooks_outliers %>%
      arrange(desc(CooksD))
  } else {
    message("当前模型未检测到显著高影响点")
  }

  
  # 计算Cook's距离并排序 (计算出实际排序)
  cooksd <- cooks.distance(销售顾问lrmod)
  n <- length(cooksd)  # 或 n <- nrow(原始数据)
  k <- length(coef(销售顾问lrmod_cooks_outliers))  # 模型参数数量
  
  threshold <- 4 / (n - k - 1)# 计算阈值
  
  # 创建数据框并筛选
  cooksd_df <- data.frame(
    Obs = 1:length(cooksd),
    CooksD = cooksd
  ) %>%
    arrange(desc(CooksD)) %>%
    filter(CooksD > threshold)
  

  
  # 输出结果
  if(nrow(cooksd_df) == 0) {
    message("无显著高影响点")
  } else {
    print(cooksd_df)
    length(cooksd_df$Obs)
  }
  

  #有147个异常值
  

  #先拿出一个看一下
  顾问销售表现数据_numeric[695,c("生效金额","预审量","进件量","活跃渠道数","拜访门店数")]

  #对比除了这个其他的
  summary(顾问销售表现数据_numeric[-614,c("生效金额","预审量","进件量","活跃渠道数","拜访门店数")])

  #所有异常值对比其他的
  outlier_index <- as.numeric(unlist(cooksd_df[,"Obs"]))

  summary(顾问销售表现数据_numeric[outlier_index, c("生效金额","预审量","进件量","活跃渠道数","拜访门店数")])  #异常值

  summary(顾问销售表现数据_numeric[-outlier_index, c("生效金额","预审量","进件量","活跃渠道数","拜访门店数")]) #剔除异常值

  #对比发现，预审的中值低于正常值，进件量中值高，拜访门店数也较高
  
  #对比一下原始值和没有异常值的数据，看看删除异常值有什么影响
  
  summary(顾问销售表现数据_numeric[, c("生效金额","预审量","进件量","活跃渠道数","拜访门店数")])#原始值
  
  #结论，对比剔除异常值的summray and 原始数据summary, 相差不大，所以剔除异常值
  
  #删除异常值
  顾问销售表现数据_numeric_v2 <- 顾问销售表现数据_numeric[ -outlier_index, ]
  
  
  #多重共线性
  
  ols_vif_tol(销售顾问lrmod)
  
  #根据经验VIF大于5或者tolerance 小于 0.2表示多重共线，需要补救，结果是所有VIF低于5，tolerance 高于0.2
  #修改的方法一般是删除那个有问题的变量，第二种是将他组合成一个变量
  
  
  #改进模型
  
  顾问销售表现数据_numeric_v2 <- 顾问销售表现数据_numeric_v2 %>% 
    mutate(预审量2 = 预审量^2) %>% 
    mutate(进件量2 = 进件量^2) %>% 
    mutate(活跃渠道数2 = 活跃渠道数^2) %>% 
    mutate(拜访门店数2 = 拜访门店数^2)
  
  销售顾问lrmod2 <- lm(
    生效金额 ~ 预审量 + 进件量 + 活跃渠道数 + 拜访门店数 +
               预审量2 + 进件量2 + 活跃渠道数2 + 拜访门店数2 
                    , 顾问销售表现数据_numeric_v2)
  
  summary(销售顾问lrmod2)
  
  #结论R方增加2%，residual standard error 减少到41.41，删除拜访门店数，活跃渠道数2，拜访门店数2
  
  销售顾问lrmod3 <- lm(
    生效金额 ~ 预审量 + 进件量 + 活跃渠道数 + 
      预审量2 + 进件量2 
    , 顾问销售表现数据_numeric_v2)

  summary(销售顾问lrmod3)
  
  # 指标	mod2	mod3	变化分析
  # R2	0.8245	0.8213	因删除3个变量，解释力略微下降（0.3%） → 可接受范围
  # 调整R² 0.8228	0.8202	调整R²下降幅度更小（0.26%），说明删除冗余变量后模型更简洁
  # F统计量 502.6	 789.4  F值大幅提升 → 模型整体显著性增强
  # 残差标准差	41.41	41.71	预测误差几乎不变（<0.7%差异） → 预测能力未受损
  
  #看来mod3是更优选，其牺牲的极少量解释力换取了：
  #更高统计效能
  #更清晰的业务洞见
  #更强的泛化能力
  
  
  #分类变量
  
  顾问销售表现数据_numeric_v2[,as.factor(c("车辆类型","渠道二级科目","店面城市"))]
  
  
  销售顾问lrmod4 <- lm(
    生效金额 ~ 预审量 + 进件量 + 活跃渠道数 + 
      预审量2 + 进件量2 + 车辆类型
    , 顾问销售表现数据_numeric_v2)
  
  summary(销售顾问lrmod4)
  
  #使用了分类变量车辆类型后，模型输出有提高到0.84

  销售顾问lrmod5 <- lm(
    生效金额 ~ 预审量 + 进件量 + 活跃渠道数 + 
      预审量2 + 进件量2 + 车辆类型 * 渠道二级科目 #(车辆类型可能和渠道二级科目交互作用)
    , 顾问销售表现数据_numeric_v2)
  
  summary(销售顾问lrmod5)
  
  #结论Residual standard error: 21.47 	Adjusted R-squared:  0.8507 又改进了一些

  
  #重要变量的选择
  
  顾问销售表现数据_numeric_v3 <- 顾问销售表现数据_numeric_v2 %>% 
    #mutate(months = as.numeric(as.numeric(生效月) - min(as.numeric(生效月)))) %>% 
    mutate(month = substr(生效月, 5, 6)) %>% 
    mutate(year = substr(生效月,0 , 4))
  
  
  ols_step_both_p(model = lm(data = 顾问销售表现数据_numeric_v3, 
                             生效金额 ~ 预审量 + 进件量 + 活跃渠道数 + 
                               预审量2 + 进件量2    + month + year),  #+"车辆类型","渠道二级科目","店面城市"
                  pent = 0.2, #变量进入阈值
                  prem = 0.01, #变量移除阈值
                  details = F)
    
    # RMSE 21.250 Adj. R-Squared 0.852 又优化了一点
  
  
  
  