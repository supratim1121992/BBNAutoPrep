#' Data inspection and quality check
#'
#' The function reads the input raw data, records summary statistics for the data and generates a report to
#' summarise data quality and characteristics.
#' @param auto Logical. If set to TRUE, the function takes in default values for certain inputs. Check the
#' scope document for the list of defaults.
#' @param prm Logical. If set to TRUE, the functions filters the data for an additional category. For the US market,
#' use prm = TRUE to filter for both Retailer and Promo Family. For the India market, can be used to filter for both
#' Brand and Cluster in case consolidated data for multiple brands is provided in a single file.
#' @return
#' @import data.table svDialogs
#' @export

Inspect_Data<-function(auto = T,prm = F){
  suppressMessages({
    suppressWarnings({
      require(svDialogs,quietly = T,warn.conflicts = F)
    })

    #####################################---Reading data---#####################################

    fl_chs<-dlgMessage(message = "Choose your data file (in .xlsx or .csv format)",type = "okcancel")$res
    if(fl_chs == "cancel"){
      stop("No data file chosen")
    }
    else if(fl_chs == "ok") {
      fl_path<-choose.files()
      wd_path<-dirname(fl_path)
      if(auto == T){
        cat(paste(wd_path,"has been set as your working directory"),
            "\nAll outputs, reports and plots will be saved here\n\n")
      }
      else {
        wd_can<-dlgMessage(message = c(paste(wd_path,"has been set as your working directory"),
                                       "All outputs, reports and plots will be saved here"),type = "ok")$res
      }

      fl_form<-unlist(strsplit(x = fl_path,split = "\\."))
      fl_form<-fl_form[length(fl_form)]
      if(auto == T){
        fl_head<-T
        fl_blnk<-c("N/A","?","-")
      }
      else {
        fl_head<-dlgMessage(message = "Does your data file contain headers/column labels?",type = "yesno")$res
        fl_head<-ifelse(test = fl_head == "yes",yes = T,no = F)
        fl_blnk<-dlgInput(message = c("Enter any characters present in the data that should be considered missing data",
                                      "For multiple inputs separate the characters by & symbol without whitespace"),default = "N/A&?&-")$res
        fl_blnk<-unlist(strsplit(fl_blnk,split = "&"))
      }
      if(fl_form == "csv"){
        dt_in<-data.table::fread(input = fl_path,sep = ",",header = fl_head,na.strings = c(fl_blnk))
      }
      else {
        if(length(readxl::excel_sheets(path = fl_path)) == 1){
          fl_sht<-1
        }
        else {
          fl_sht<-dlgList(choices = readxl::excel_sheets(path = fl_path),multiple = F,title = "Select sheet")$res
        }
        dt_in<-as.data.table(readxl::read_excel(path = fl_path,sheet = fl_sht,col_names = fl_head,na = c(""," ",fl_blnk)))
      }
      if(exists("dt_in") == T){
        col_ret<-dlgList(choices = colnames(dt_in),title = "Choose Market/Retailer column")$res
        if(prm == T){
          col_prm<-dlgList(choices = colnames(dt_in),title = "Choose Promo family column")$res
        }

        sel_yr<-dlgList(choices = colnames(dt_in),title = "Choose year column")$res
        sel_per<-dlgList(choices = colnames(dt_in),title = "Choose period/month column")$res
        colnames(dt_in)[which(colnames(dt_in) %in% col_ret)]<-"Market"
        if(prm == T){
          colnames(dt_in)[which(colnames(dt_in) %in% col_prm)]<-"Promo"
        }
        dt_in[,Year_Temp := as.numeric(dt_in[[sel_yr]])]
        dt_in[,Period_Temp := as.numeric(gsub(pattern = '([0-9]+).*',replacement = '\\1',x = dt_in[[sel_per]]))]
        if(prm == T){
          dt_in<-dt_in[order(Market,Promo,Year_Temp,Period_Temp,decreasing = F)]
          colnames(dt_in)[which(colnames(dt_in) %in% "Promo")]<-col_prm
        }
        else {
          dt_in<-dt_in[order(Market,Year_Temp,Period_Temp,decreasing = F)]
          colnames(dt_in)[which(colnames(dt_in) %in% "Market")]<-col_ret
        }
        dt_in<-dt_in[,Year_Temp := NULL]
        dt_in<-dt_in[,Period_Temp := NULL]


        Check_cls<-function(x){
          x_na<-na.omit(x)
          x_num<-as.numeric(x_na)
          if(sum(is.na(x_num)) > 0.5 * length(x_na)){
            is_num<-F
          }
          else {
            is_num<-T
          }
          return(is_num)
        }

        suppressWarnings({
          vec_isnum<-as.logical(dt_in[,lapply(X = .SD,FUN = Check_cls),.SDcols = colnames(dt_in)][1,])
          names(vec_isnum)<-colnames(dt_in)
          vec_isnum[which(names(vec_isnum) %in% c(sel_yr,sel_per))]<-F
          dt_in_num<-dt_in[,lapply(X = .SD,FUN = as.numeric),.SDcols = colnames(dt_in)[vec_isnum]]
          dt_in_cat<-dt_in[,!vec_isnum,with = F]
          for(i in 1:ncol(dt_in_num)){
            dt_in[,which(colnames(dt_in) == colnames(dt_in_num)[i])]<-dt_in_num[[i]]
          }
        })
        if(auto == T){
          cat("-----Data successfully read-----\n\n")
        }
        else {
          dlgMessage(message = "Data successfully read",type = "ok")
          dt_prv<-dlgMessage(message = c(paste("No. of observations (rows) in the dataset :",nrow(dt_in)),
                                         paste("No. of variables (columns) in the dataset :",ncol(dt_in)),
                                         "Would you like to preview a portion of your data?"),type = "yesno")$res
          if(dt_prv == "yes"){
            print(dt_in[1:10,1:10])
          }
        }
      }
      else {
        stop("The data could not be read. Please make sure the format of the file is supported.")
      }
    }

    #####################################---Report generation---#####################################

    fl_nm_ls<-unlist(strsplit(x = fl_path,split = "\\",fixed = T))
    fl_nm<-fl_nm_ls[length(fl_nm_ls)]
    if(exists("fl_sht")){
      fl_nm<-paste(fl_nm,fl_sht,sep = " - Sheet: ")
    }
    fl_size<-file.info(fl_path)$size/1000
    fl_src<-dlgInput(message = "Enter file source")$res
    fl_dt_run<-dlgInput(message = c("Enter file received date","Format : YYYY-MM-DD"),
                        default = "2019-01-31")$res
    fl_dt_exp<-dlgInput(message = c("Enter expected date","Format : YYYY-MM-DD"),
                        default = "2019-01-31")$res
    dt_fl_sum<-data.table("File name" = fl_nm,"File path" = fl_path,"Size (KB)" = fl_size,
                          "Variables" = ncol(dt_in),"Observations" = nrow(dt_in),
                          "Data source" = fl_src,"Date received" = fl_dt_run,"Date expected" = fl_dt_exp)

    cat("GENERATING REPORTS\n")
    if(prm == T){
      dt_in_mix<-cbind.data.frame("Market_Retailer" = dt_in[[col_ret]],"Promo_Family" = dt_in[[col_prm]],dt_in_num)
      dt_in_cix<-cbind.data.frame("Market_Retailer" = dt_in[[col_ret]],"Promo_Family" = dt_in[[col_prm]],dt_in_cat)
      vec_prm<-unique(na.omit(dt_in[[col_prm]]))
    }
    else {
      dt_in_mix<-cbind.data.frame("Market_Retailer" = dt_in[[col_ret]],dt_in_num)
      dt_in_cix<-cbind.data.frame("Market_Retailer" = dt_in[[col_ret]],dt_in_cat)
    }
    vec_ret<-unique(na.omit(dt_in[[col_ret]]))
    ls_dt_num<-vector(mode = "list",length = length(vec_ret))
    ls_dt_out<-vector(mode = "list",length = length(vec_ret))
    ls_dt_cat<-vector(mode = "list",length = length(vec_ret))
    names(ls_dt_num)<-names(ls_dt_out)<-names(ls_dt_cat)<-vec_ret

    cat("|",rep(x = "=",times = 20),sep = "")

    for(i in 1:length(vec_ret)){
      if(prm == T){
        for(j in 1:length(vec_prm)){
          ls_dt_num[[i]][[j]]<-dt_in_mix[`Market_Retailer` == vec_ret[i] & `Promo_Family` == vec_prm[j]]
          names(ls_dt_num[[i]])[j]<-paste(vec_ret[i],vec_prm[j],sep = "_")
          num_id<-which(dt_in_mix$`Market_Retailer` == vec_ret[i] & dt_in_mix$`Promo_Family` == vec_prm[j])
          ls_out<-lapply(X = ls_dt_num[[i]][[j]][,-c(1,2),with = F],FUN = function(x){
            x_na<-na.omit(x)
            x_val<-x_na[which(x_na > 0)]
            x_quant<-quantile(x = x_val,na.rm = T)
            x_iqr<-IQR(x = x_val,na.rm = T)
            x_low<-which(is.na(x) == F & x > 0 & x < (x_quant[2] - 1.5*x_iqr))
            x_up<-which(is.na(x) == F & x > 0 & x > (x_quant[4] + 1.5*x_iqr))
            outs<-sort(x = c(x_low,x_up),decreasing = F)
            outs<-num_id[outs]
            return(outs)
          })
          ls_dt_out[[i]][[j]]<-ls_out
          names(ls_dt_out[[i]])[j]<-paste(vec_ret[i],vec_prm[j],sep = "_")
          ls_dt_cat[[i]][[j]]<-dt_in_cix[`Market_Retailer` == vec_ret[i] & `Promo_Family` == vec_prm[j]]
          names(ls_dt_cat[[i]])[j]<-paste(vec_ret[i],vec_prm[j],sep = "_")
        }
      }

      else {
        ls_dt_num[[i]]<-dt_in_mix[`Market_Retailer` == vec_ret[i]]
        num_id<-which(dt_in_mix$`Market_Retailer` == vec_ret[i])
        ls_out<-lapply(X = ls_dt_num[[i]][,-1,with = F],FUN = function(x){
          x_na<-na.omit(x)
          x_val<-x_na[which(x_na > 0)]
          x_quant<-quantile(x = x_val,na.rm = T)
          x_iqr<-IQR(x = x_val,na.rm = T)
          x_low<-which(is.na(x) == F & x > 0 & x < (x_quant[2] - 1.5*x_iqr))
          x_up<-which(is.na(x) == F & x > 0 & x > (x_quant[4] + 1.5*x_iqr))
          outs<-sort(x = c(x_low,x_up),decreasing = F)
          outs<-num_id[outs]
          return(outs)
        })
        ls_dt_out[[i]]<-ls_out
        ls_dt_cat[[i]]<-dt_in_cix[`Market_Retailer` == vec_ret[i]]
      }
    }

    cat(rep(x = "=",times = 20),sep = "")

    Create_num_sum<-function(y){
      mkt<-unique(na.omit(y[[1]]))
      if(prm == T){
        pfm<-unique(na.omit(y[[2]]))
        dt_nrt<-y[,-c(1,2),with = F]
      }
      else {
        dt_nrt<-y[,-1,with = F]
      }
      nm_sum<-as.data.table(lapply(X = dt_nrt,FUN = function(x){
        return(summary(x,na.rm = T)[1:6])
      }))
      nm_sum<-nm_sum[,lapply(X = .SD,FUN = round,digits = 2),.SDcols = colnames(nm_sum)]
      var_sum<-colnames(nm_sum)
      mkt_col<-rep(x = mkt,times = length(var_sum))
      if(prm == T){
        pfm_col<-rep(x = pfm,times = length(var_sum))
        nm_sum<-cbind.data.frame(var_sum,mkt_col,pfm_col,transpose(nm_sum))
        colnames(nm_sum)<-c("Variable","Market_Retailer","Promo_Family","Min","Q1","Median","Mean","Q3","Max")
      }
      else {
        nm_sum<-cbind.data.frame(var_sum,mkt_col,transpose(nm_sum))
        colnames(nm_sum)<-c("Variable","Market_Retailer","Min","Q1","Median","Mean","Q3","Max")
      }
      nm_sum[,Total := nrow(dt_nrt)]
      na_cnt<-as.numeric(dt_nrt[,lapply(X = .SD,FUN = function(x){
        return(sum(is.na(x)))
      }),.SDcols = colnames(dt_nrt)][1,])
      zero_cnt<-as.numeric(dt_nrt[,lapply(X = .SD,FUN = function(x){
        return(length(which(x == 0)))
      }),.SDcols = colnames(dt_nrt)][1,])
      neg_cnt<-as.numeric(dt_nrt[,lapply(X = .SD,FUN = function(x){
        return(length(which(x < 0)))
      }),.SDcols = colnames(dt_nrt)][1,])
      out_cnt<-as.numeric(dt_nrt[,lapply(X = .SD,FUN = function(x){
        x<-na.omit(x)
        x_val<-x[which(x > 0)]
        x_quant<-quantile(x = x_val,na.rm = T)
        x_low<-which(x_val < (x_quant[2] - 1.5*IQR(x = x_val,na.rm = T)))
        x_up<-which(x_val > (x_quant[4] + 1.5*IQR(x = x_val,na.rm = T)))
        outs<-length(x_low) + length(x_up)
        return(outs)
      }),.SDcols = colnames(dt_nrt)][1,])
      nm_sum<-cbind.data.frame(nm_sum,"Missing" = na_cnt,"Zeros" = zero_cnt,"Negative" = neg_cnt,
                               "Outliers" = out_cnt)
      nm_sum[,Valid := Total - Missing - Zeros - Negative - Outliers]
      nm_sum[,"Missing(%)" := round(Missing/Total * 100,digits = 2)]
      nm_sum[,"Zeros(%)" := round(Zeros/Total * 100,digits = 2)]
      nm_sum[,"Negative(%)" := round(Negative/Total * 100,digits = 2)]
      nm_sum[,"Outliers(%)" := round(Outliers/Total * 100,digits = 2)]
      nm_sum[,"Valid(%)" := round(Valid/Total * 100,digits = 2)]

      return(nm_sum)
    }

    if(prm == T){
      ls_dt_num<-lapply(X = ls_dt_num,FUN = function(y){
        lapply(X = y,FUN = Create_num_sum)
      })
      num_sum<-rbindlist(l = lapply(X = ls_dt_num,FUN = rbindlist,use.names = T),use.names = T)
    }
    else {
      ls_dt_num<-lapply(X = ls_dt_num,FUN = Create_num_sum)
      num_sum<-rbindlist(l = ls_dt_num,use.names = T)
    }
    num_sum<-num_sum[,lapply(X = .SD,FUN = function(x){
      x[is.nan(x)]<-NA
      return(x)
    }),.SDcols = colnames(num_sum)]
    num_sum<-num_sum[order(`Valid(%)`)]

    Create_cat_sum<-function(y){
      if(prm == T){
        mode_cat<-as.character(y[,lapply(X = .SD,FUN = function(x){
          x_tab<-table(x,useNA = "no")
          mode<-names(x_tab)[which.max(x_tab)]
        }),.SDcols = colnames(y)[-c(1,2)]][1,])
        lvl_cnt<-as.numeric(y[,lapply(X = .SD,FUN = function(x){
          return(length(na.omit(unique(x))))
        }),.SDcols = colnames(y)[-c(1,2)]][1,])
        mod_cnt<-as.numeric(y[,lapply(X = .SD,FUN = function(x){
          return(max(table(x,useNA = "no")))
        }),.SDcols = colnames(y)[-c(1,2)]][1,])
        nac_cnt<-as.numeric(y[,lapply(X = .SD,FUN = function(x){
          return(sum(is.na(x)))
        }),.SDcols = colnames(y)[-c(1,2)]][1,])
        mkt<-unique(na.omit(y[[1]]))
        col_mkt<-rep(x = mkt,times = length(mode_cat))
        pfm<-unique(na.omit(y[[2]]))
        col_pfm<-rep(x = pfm,times = length(mode_cat))
        cat_sum<-data.table("Variable" = colnames(y)[-c(1,2)],"Market_Retailer" = col_mkt,"Promo_Family" = col_pfm,
                            "Unique levels" = lvl_cnt,"Mode" = mode_cat,"Mode count" = mod_cnt,
                            "Missing" = nac_cnt,"Total" = rep(x = nrow(y),times = length(colnames(y)) - 2))
        cat_sum[,"Mode count(%)" := round(`Mode count`/Total * 100,digits = 2)]
        cat_sum[,"Missing(%)" := round(Missing/Total * 100,digits = 2)]
      }
      else {
        mode_cat<-as.character(y[,lapply(X = .SD,FUN = function(x){
          x_tab<-table(x,useNA = "no")
          mode<-names(x_tab)[which.max(x_tab)]
        }),.SDcols = colnames(y)[-1]][1,])
        lvl_cnt<-as.numeric(y[,lapply(X = .SD,FUN = function(x){
          return(length(na.omit(unique(x))))
        }),.SDcols = colnames(y)[-1]][1,])
        mod_cnt<-as.numeric(y[,lapply(X = .SD,FUN = function(x){
          return(max(table(x,useNA = "no")))
        }),.SDcols = colnames(y)[-1]][1,])
        nac_cnt<-as.numeric(y[,lapply(X = .SD,FUN = function(x){
          return(sum(is.na(x)))
        }),.SDcols = colnames(y)[-1]][1,])
        mkt<-unique(na.omit(y[[1]]))
        col_mkt<-rep(x = mkt,times = length(mode_cat))
        cat_sum<-data.table("Variable" = colnames(y)[-1],"Market_Retailer" = col_mkt,"Unique levels" = lvl_cnt,
                            "Mode" = mode_cat,"Mode count" = mod_cnt,"Missing" = nac_cnt,
                            "Total" = rep(x = nrow(y),times = length(colnames(y)[-1])))
        cat_sum[,"Mode count(%)" := round(`Mode count`/Total * 100,digits = 2)]
        cat_sum[,"Missing(%)" := round(Missing/Total * 100,digits = 2)]
      }

      return(cat_sum)
    }

    if(prm == T){
      ls_dt_cat<-lapply(X = ls_dt_cat,FUN = function(y){
        lapply(X = y,FUN = Create_cat_sum)
      })
      cat_sum<-rbindlist(l = lapply(X = ls_dt_cat,FUN = rbindlist,use.names = T),use.names = T)
    }
    else {
      ls_dt_cat<-lapply(X = ls_dt_cat,FUN = Create_cat_sum)
      cat_sum<-rbindlist(l = ls_dt_cat,use.names = T)
    }
    cat_sum<-cat_sum[order(-`Missing(%)`)]

    ls_mis<-lapply(X = dt_in,FUN = function(x){
      which(is.na(x) == T)
    })
    ls_mis<-ls_mis[lapply(X = ls_mis,FUN = length) > 0]
    ls_neg<-lapply(X = dt_in_num,FUN = function(x){
      which(x < 0)
    })
    ls_neg<-ls_neg[lapply(X = ls_neg,FUN = length) > 0]
    ls_zero<-lapply(X = dt_in_num,FUN = function(x){
      which(x == 0)
    })
    ls_zero<-ls_zero[lapply(X = ls_zero,FUN = length) > 0]
    ls_out<-vector(mode = "list",length = length(colnames(dt_in_num)))
    names(ls_out)<-colnames(dt_in_num)
    if(prm == T){
      for(i in 1:length(vec_ret)){
        for(j in 1:length(vec_prm)){
          ls_out<-mapply(FUN = c,ls_out,ls_dt_out[[i]][[j]])
        }
      }
    }
    else {
      for(i in 1:length(ls_dt_out)){
        ls_out<-mapply(c,ls_out,ls_dt_out[[i]])
      }
    }
    ls_out<-ls_out[lapply(X = ls_out,FUN = length) > 0]

    cat(rep(x = "=",times = 20),sep = "")

    wb_rep<-openxlsx::createWorkbook()
    hed_style<-openxlsx::createStyle(border = c("top","bottom","left","right"),borderStyle = "medium",halign = "center",
                                     textDecoration = "bold")
    openxlsx::addWorksheet(wb = wb_rep,sheetName = "Data info",gridLines = T)
    openxlsx::writeData(wb = wb_rep,sheet = "Data info",x = dt_fl_sum,colNames = T,rowNames = F,headerStyle = hed_style,
                        borders = "all",borderStyle = "medium")
    openxlsx::setColWidths(wb = wb_rep,sheet = "Data info",cols = 1:ncol(dt_fl_sum),widths = "auto")
    openxlsx::addWorksheet(wb = wb_rep,sheetName = "Summary (numeric)",gridLines = T)
    openxlsx::writeDataTable(wb = wb_rep,sheet = "Summary (numeric)",x = num_sum,colNames = T,rowNames = F)
    openxlsx::setColWidths(wb = wb_rep,sheet = "Summary (numeric)",cols = 1:ncol(num_sum),widths = "auto")
    openxlsx::addWorksheet(wb = wb_rep,sheetName = "Summary (categorical)",gridLines = T)
    openxlsx::writeDataTable(wb = wb_rep,sheet = "Summary (categorical)",x = cat_sum,colNames = T,rowNames = F)
    openxlsx::setColWidths(wb = wb_rep,sheet = "Summary (categorical)",cols = 1:ncol(cat_sum),widths = "auto")
    openxlsx::addWorksheet(wb = wb_rep,sheetName = "Data highlight",gridLines = T)
    openxlsx::writeDataTable(wb = wb_rep,sheet = "Data highlight",x = dt_in,colNames = T,rowNames = F)
    openxlsx::setColWidths(wb = wb_rep,sheet = "Data highlight",cols = 1:ncol(dt_in),widths = "auto")
    mis_style<-openxlsx::createStyle(fgFill = "#b8d8e1")
    out_style<-openxlsx::createStyle(fgFill = "#357388")
    neg_style<-openxlsx::createStyle(fgFill = "#528c9e")
    zero_style<-openxlsx::createStyle(fgFill = "#7fb4be")

    if(length(ls_mis) > 0){
      for(i in 1:length(ls_mis)){
        openxlsx::addStyle(wb = wb_rep,sheet = "Data highlight",style = mis_style,cols = which(colnames(dt_in) %in% names(ls_mis)[i]),
                           rows = ls_mis[[i]] + 1)
      }
    }
    if(length(ls_neg) > 0){
      for(i in 1:length(ls_neg)){
        openxlsx::addStyle(wb = wb_rep,sheet = "Data highlight",style = neg_style,cols = which(colnames(dt_in) %in% names(ls_neg)[i]),
                           rows = ls_neg[[i]] + 1)
      }
    }
    if(length(ls_zero) > 0){
      for(i in 1:length(ls_zero)){
        openxlsx::addStyle(wb = wb_rep,sheet = "Data highlight",style = zero_style,cols = which(colnames(dt_in) %in% names(ls_zero)[i]),
                           rows = ls_zero[[i]] + 1)
      }
    }
    if(length(ls_out) > 0){
      for(i in 1:length(ls_out)){
        openxlsx::addStyle(wb = wb_rep,sheet = "Data highlight",style = out_style,cols = which(colnames(dt_in) %in% names(ls_out)[i]),
                           rows = ls_out[[i]] + 1)
      }
    }
    low_style<-openxlsx::createStyle(fgFill = "#C2F2B5")
    med_style<-openxlsx::createStyle(fgFill = "#F6EDA0")
    high_style<-openxlsx::createStyle(fgFill = "#EFAFBF")
    num_low<-which(num_sum$`Valid(%)` >= 75)
    num_med<-which(num_sum$`Valid(%)` >= 50 & num_sum$`Valid(%)` < 75)
    num_high<-which(num_sum$`Valid(%)` < 50)
    cat_high<-which(cat_sum$`Missing(%)` >= 75)
    cat_med<-which(cat_sum$`Missing(%)` >= 50 & cat_sum$`Missing(%)` < 75)
    cat_low<-which(cat_sum$`Missing(%)` < 50)

    openxlsx::addStyle(wb = wb_rep,sheet = "Summary (numeric)",style = low_style,rows = num_low + 1,cols = 1:ncol(num_sum),gridExpand = T)
    openxlsx::addStyle(wb = wb_rep,sheet = "Summary (numeric)",style = med_style,rows = num_med + 1,cols = 1:ncol(num_sum),gridExpand = T)
    openxlsx::addStyle(wb = wb_rep,sheet = "Summary (numeric)",style = high_style,rows = num_high + 1,cols = 1:ncol(num_sum),gridExpand = T)
    openxlsx::addStyle(wb = wb_rep,sheet = "Summary (categorical)",style = low_style,rows = cat_low + 1,cols = 1:ncol(cat_sum),gridExpand = T)
    openxlsx::addStyle(wb = wb_rep,sheet = "Summary (categorical)",style = med_style,rows = cat_med + 1,cols = 1:ncol(cat_sum),gridExpand = T)
    openxlsx::addStyle(wb = wb_rep,sheet = "Summary (categorical)",style = high_style,rows = cat_high + 1,cols = 1:ncol(cat_sum),gridExpand = T)

    openxlsx::saveWorkbook(wb = wb_rep,file = paste(wd_path,paste(paste("Data inspection",unlist(strsplit(x = fl_nm_ls[length(fl_nm_ls)],split = "\\."))[1],sep = "_")
                                                                  ,".xlsx",sep = ""),sep = "\\"),overwrite = T)

    cat(rep(x = "=",times = 20),"|\n\n",sep = "")

    dlgMessage(message = "Data summary report has been successfully generated and saved in your working directory")
  })
}

#' Automated data processing and preparation for BBN model creation
#'
#' The function reads the input raw data, provides the user a set of options to treat the data and writes out an
#' aggregated data with all the variables treated.
#' @param prm Logical. If set to TRUE, the functions filters the data for an additional category. For the US market,
#' use prm = TRUE to filter for both Retailer and Promo Family. For the India market, can be used to filter for both
#' Brand and Cluster in case consolidated data for multiple brands is provided in a single file.
#' @return A list containing the aggregated processed data, the time period and scope document information along
#' with the path of the working directory. These are read as input in the Transform_Data function.
#' @import data.table svDialogs
#' @export

Prepare_Data<-function(prm = F){
  suppressMessages({
    suppressWarnings({
      require(svDialogs,quietly = T,warn.conflicts = F)
    })

    #####################################---Reading data---#####################################

    fl_chs<-dlgMessage(message = "Choose your data file (in .xlsx or .csv format)",type = "okcancel")$res
    if(fl_chs == "cancel"){
      stop("No data file chosen")
    }
    else if(fl_chs == "ok") {
      fl_path<-choose.files()
      wd_path<-dirname(fl_path)
      cat(paste(wd_path,"has been set as your working directory"),
          "\nAll outputs, reports and plots will be saved here\n\n")

      fl_form<-unlist(strsplit(x = fl_path,split = "\\."))
      fl_form<-fl_form[length(fl_form)]
      fl_head<-T
      fl_blnk<-c("N/A","?","-","NA")

      if(fl_form == "csv"){
        dt_in<-data.table::fread(input = fl_path,sep = ",",header = fl_head,na.strings = c(fl_blnk))
      }
      else {
        if(length(readxl::excel_sheets(path = fl_path)) == 1){
          fl_sht<-1
        }
        else {
          fl_sht<-dlgList(choices = readxl::excel_sheets(path = fl_path),multiple = F,title = "Select sheet")$res
        }
        dt_in<-as.data.table(readxl::read_excel(path = fl_path,sheet = fl_sht,col_names = fl_head,na = c(""," ",fl_blnk)))
      }
      if(exists("dt_in") == T){
        Check_cls<-function(x){
          x_na<-na.omit(x)
          x_num<-as.numeric(x_na)
          if(sum(is.na(x_num)) > 0.5 * length(x_na)){
            is_num<-F
          }
          else {
            is_num<-T
          }
          return(is_num)
        }

        suppressWarnings({
          vec_isnum<-as.logical(dt_in[,lapply(X = .SD,FUN = Check_cls),.SDcols = colnames(dt_in)][1,])
          dt_in_num<-dt_in[,lapply(X = .SD,FUN = as.numeric),.SDcols = colnames(dt_in)[vec_isnum]]
          for(i in 1:ncol(dt_in_num)){
            dt_in[,which(colnames(dt_in) == colnames(dt_in_num)[i])]<-dt_in_num[[i]]
          }
        })

        cat("-----Data successfully read-----\n\n")
      }
      else {
        stop("The data could not be read. Please make sure the format of the file is supported.")
      }
    }

    #####################################---Data subsetting---#####################################

    dlgMessage(message = "Choose Scope document / Master input file",type = "ok")
    scp_path<-choose.files()
    dt_scope<-as.data.table(readxl::read_excel(path = scp_path,sheet = "Scope",col_names = T))
    dt_sign<-as.data.table(readxl::read_excel(path = scp_path,sheet = "Signs",col_names = T))
    dt_scp<-list("Scope" = dt_scope,"Sign" = dt_sign)

    col_ret<-dlgList(choices = colnames(dt_in),title = "Choose Market/Retailer column")$res
    unq_ret<-unique(na.omit(dt_in[[col_ret]]))
    if(prm == T){
      col_prm<-dlgList(choices = colnames(dt_in),title = "Choose promo family column")$res
    }

    sel_yr<-dlgList(choices = colnames(dt_in),title = "Choose year column")$res
    sel_per<-dlgList(choices = colnames(dt_in),title = "Choose period/month column")$res
    dt_in[,Year_Temp := as.numeric(dt_in[[sel_yr]])]
    dt_in[,Period_Temp := as.numeric(gsub(pattern = '([0-9]+).*',replacement = '\\1',x = dt_in[[sel_per]]))]
    dt_in<-dt_in[order(Year_Temp,Period_Temp)]
    vec_time<-unique(paste(dt_in$Year_Temp,paste("Period",dt_in$Period_Temp,sep = ":"),sep = " - "))
    dt_in<-dt_in[,Year_Temp := NULL]
    dt_in<-dt_in[,Period_Temp := NULL]

    seas_op<-dlgMessage(message = "Calculate seasonality index?",type = "yesno")$res
    if(seas_op == "yes"){
      seas_nm<-dt_scp$Scope[["Model_Name"]][which(dt_scp$Scope[["RawData_Name"]] %in% colnames(dt_in))]
      seas_var<-dlgList(choices = seas_nm,multiple = F,title = "Choose variable for Seas.ID")$res
      seas_var<-dt_scp$Scope[["RawData_Name"]][which(dt_scp$Scope[["Model_Name"]] %in% seas_var)]
    }
    cat_op<-dlgMessage(message = "Calculate category index?",type = "yesno")$res
    if(cat_op == "yes"){
      cat_nm<-dt_scp$Scope[["Model_Name"]][which(dt_scp$Scope[["RawData_Name"]]%in% colnames(dt_in))]
      cat_var<-dlgList(choices = cat_nm,multiple = F,title = "Choose variable for Cat.ID")$res
      cat_var<-dt_scp$Scope[["RawData_Name"]][which(dt_scp$Scope[["Model_Name"]] %in% cat_var)]
    }

    ls_dt_in<-vector(mode = "list",length = length(unq_ret))
    names(ls_dt_in)<-unq_ret
    for(i in 1:length(unq_ret)){
      ls_dt_in[[i]]<-dt_in[eval(as.name(col_ret)) %in% unq_ret[i]]
    }
    if(exists("col_prm")){
      data_slice<-function(x,y){
        unq_prm<-unique(na.omit(x[[y]]))
        ls_dt_slc<-vector(mode = "list",length = length(unq_prm))
        names(ls_dt_slc)<-unq_prm
        for(i in 1:length(unq_prm)){
          ls_dt_slc[[i]]<-x[eval(as.name(y)) %in% unq_prm[i]]
        }
        return(ls_dt_slc)
      }
      ls_dt_in<-lapply(X = ls_dt_in,FUN = data_slice,y = col_prm)
    }

    set_time<-function(x){
      x[,Year_Temp := as.numeric(x[[sel_yr]])]
      x[,Period_Temp := as.numeric(gsub(pattern = '([0-9]+).*',replacement = '\\1',x = x[[sel_per]]))]
      if(seas_op == "yes"){
        dt_seas<-x[,.("seasonality.index" = mean(as.numeric(eval(as.name(seas_var))),na.rm = T)),by = "Period_Temp"]
        dt_seas[,seasonality.index := seasonality.index/mean(seasonality.index,na.rm = T)]
        x<-merge(x = x,y = dt_seas,by = "Period_Temp")
      }
      if(cat_op == "yes"){
        dt_cat<-x[,.("category.index" = mean(as.numeric(eval(as.name(cat_var))),na.rm = T)),by = "Period_Temp"]
        dt_cat[,category.index := category.index/mean(category.index,na.rm = T)]
        x<-merge(x = x,y = dt_cat,by = "Period_Temp")
      }
      x<-x[order(Year_Temp,Period_Temp)]
      x[,Year_Temp := NULL]
      x[,Period_Temp := NULL]

      return(x)
    }
    if(prm == T){
      ls_dt_in<-lapply(X = ls_dt_in,FUN = function(y){
        lapply(X = y,FUN = set_time)
      })
    }
    else {
      ls_dt_in<-lapply(X = ls_dt_in,FUN = set_time)
    }

    if(seas_op == "yes"){
      cat("-----Seasonality index successfully computed-----\n\n")
    }
    if(cat_op == "yes"){
      cat("-----Category index successfully computed-----\n\n")
    }

    vec_var_keep<-dt_scp$Scope[["RawData_Name"]]

    sel_col<-function(x,ref_keep){
      x<-x[,which(colnames(x) %in% ref_keep == T),with = F]
      x<-x[,ref_keep[which(ref_keep %in% colnames(x) == T)],with = F]
      return(x)
    }
    if(prm == T){
      ls_dt_sub<-lapply(X = ls_dt_in,FUN = function(y){
        lapply(X = y,FUN = sel_col,ref_keep = vec_var_keep)
      })
      ls_dt_sub<-lapply(X = ls_dt_sub,FUN = function(y){
        lapply(X = y,FUN = function(x){
          suppressWarnings({
            x[,lapply(X = .SD,FUN = as.numeric),.SDcols = colnames(x)]
          })
          return(x)
        })
      })
    }
    else {
      ls_dt_sub<-lapply(X = ls_dt_in,FUN = sel_col,ref_keep = vec_var_keep)
      ls_dt_sub<-lapply(X = ls_dt_sub,FUN = function(x){
        suppressWarnings({
          x[,lapply(X = .SD,FUN = as.numeric),.SDcols = colnames(x)]
        })
        return(x)
      })
    }

    cat("-----In scope variables selected-----\n\n")

    rename<-function(y){
      new_nm<-sapply(X = colnames(y),FUN = function(x){
        x<-dt_scp$Scope[["Model_Name"]][which(dt_scp$Scope[["RawData_Name"]] %in% x)]
        return(x)
      })
      colnames(y)<-new_nm
      return(y)
    }

    if(prm == T){
      ls_dt_sub<-lapply(X = ls_dt_sub,FUN = function(y){
        lapply(X = y,FUN = rename)
      })
      dt_log_trt<-data.table("Market/Retailer" = character(),"Promo" = character(),"Variable" = character(),
                             "Period" = character(),"Treatment" = character(),"Treatment technique" = character(),
                             "Actual" = numeric(),"Treated" = numeric())
    }
    else {
      ls_dt_sub<-lapply(X = ls_dt_sub,FUN = rename)
      dt_log_trt<-data.table("Market/Retailer" = character(),"Variable" = character(),"Period" = character(),
                             "Treatment" = character(),"Treatment technique" = character(),
                             "Actual" = numeric(),"Treated" = numeric())
    }

    cat("-----Column renaming complete------\n\n")

    #####################################---Missing value treatment---#####################################

    if(anyNA(ls_dt_sub,recursive = T) == T){
      imp<-"yes"
      chk_imp<-dlgMessage(message = c("Missing values have been detected in the data",
                                      "Follow the upcoming prompts for imputation"),type = "okcancel")$res
      if(chk_imp == "cancel"){
        stop("The function execution has been stopped")
      }

      treat_NA<-function(x,name){
        rep_na<-x[,lapply(X = .SD,FUN = function(x){
          sum(is.na(x))
        }),.SDcols = colnames(x)]
        na_id<-which(as.numeric(rep_na[1,]) > 0)
        mis_col<-colnames(x)[which(as.numeric(rep_na[1,]) == nrow(x))]

        dt_log_mis<-data.table("Variable" = character(),"Period" = character(),"Treatment" = character(),
                               "Treatment technique" = character(),"Actual" = numeric(),"Treated" = numeric())

        if(length(mis_col) > 0){
          na_chk<-dlgMessage(message = c("Variables with missing values throughout have been detected",
                                         "Replace them with a fixed value to avoid errors in the module"),type = "okcancel")$res
          if(na_chk == "cancel"){
            stop("Function execution has been stopped")
          }
          val<-as.numeric(dlgInput(message = "Enter the value (numeric) to replace missing values",default = 0)$res)
          for(i in 1:length(mis_col)){
            na_row<-which(is.na(x[[which(colnames(x) %in% mis_col[i])]]) == T)
            log_na_act<-x[[mis_col[i]]][na_row]
            x[na_row,which(colnames(x) %in% mis_col[i])]<-val
            log_na_mod<-x[[mis_col[i]]][na_row]
            dt_log_mis_fix<-data.table("Variable" = rep(x = mis_col[i],times = length(na_row)),
                                       "Period" = vec_time[na_row],"Treatment" = rep(x = "Missing",times = length(na_row)),
                                       "Treatment technique" = rep(x = "Fixed value",times = length(na_row)),
                                       "Actual" = log_na_act,"Treated" = log_na_mod)
            dt_log_mis<-rbind.data.frame(dt_log_mis,dt_log_mis_fix)
          }
          cat("Variables missing throughout have been treated\n\n")
        }

        if(sum(complete.cases(x)) > 15){
          imp_choice<-c("knnImputation","Impute with median of same year","Impute with mean of same year","Imputation with Rolling mean","Impute with fixed value")
        }
        else {
          imp_choice<-c("Impute with median of same year","Impute with mean of same year","Imputation with Rolling mean","Impute with fixed value")
        }
        while(sum(is.na(x)) > 0 & imp == "yes"){
          dt_na<-x[,lapply(X = .SD,FUN = function(x){
            sum(is.na(x))
          }),.SDcols = colnames(x)]
          na_vec<-as.integer(dt_na[1,])
          na_col<-which(na_vec > 0)
          dt_na<-dt_na[,na_col,with = F]
          na_var<-character(length = ncol(dt_na))
          for(i in 1:ncol(dt_na)){
            na_pct<-round(dt_na[[i]][1]/nrow(x) * 100,digits = 2)
            na_var[i]<-paste(colnames(dt_na)[i],paste(na_pct,"% missing",sep = ""),paste(dt_na[[i]][1],"values",sep = " "),sep = " ~ ")
          }
          imp_sel<-dlgList(choices = imp_choice,title = "Choose imputation technique")$res
          if(imp_sel == "knnImputation"){
            ls_na_row<-lapply(X = x,FUN = function(y){
              na_row<-which(is.na(y) == T)
              log_na_act<-y[na_row]

              return(list("NA row" = na_row,"Actual" = log_na_act))
            })
            x<-DMwR::knnImputation(x)
            for(i in 1:length(ls_na_row)){
              log_na_mod<-x[[i]][ls_na_row[[i]]$`NA row`]
              dt_log_mis_knn<-data.table("Variable" = rep(x = names(ls_na_row)[i],times = length(ls_na_row[[i]]$`NA row`)),
                                         "Period" = vec_time[ls_na_row[[i]]$`NA row`],"Treatment" = rep(x = "Missing",times = length(ls_na_row[[i]]$`NA row`)),
                                         "Treatment technique" = rep(x = "KNN",times = length(ls_na_row[[i]]$`NA row`)),
                                         "Actual" = ls_na_row[[i]]$`Actual`,"Treated" = log_na_mod)
              dt_log_mis<-rbind.data.frame(dt_log_mis,dt_log_mis_knn)
            }
          }
          else if(imp_sel == "Impute with median of same year"){
            imp_med<-dlgList(choices = c(na_var,"---None---"),title = "Choose the variables to impute with median",multiple = T)$res
            if("---None---" %in% imp_med == F){
              imp_med<-unlist(lapply(strsplit(x = imp_med,split = " ~ ",fixed = T),function(x){x[[1]]}))
              vec_year<-unlist(lapply(strsplit(x = vec_time,split = " - ",fixed = T),function(x){x[[1]]}))
              for(i in 1:length(imp_med)){
                na_row<-which(is.na(x[[which(colnames(x) %in% imp_med[i])]]) == T)
                na_yr<-vec_year[na_row]
                log_na_act<-x[[imp_med[i]]][na_row]
                for(j in 1:length(unique(na_yr))){
                  if(length(x[[which(colnames(x) %in% imp_med[i])]][which(vec_year %in% unique(na_yr)[j])]) == sum(is.na(x[[which(colnames(x) %in% imp_med[i])]][which(vec_year %in% unique(na_yr)[j])]))){
                    x[which(is.na(x[[which(colnames(x) %in% imp_med[i])]]) == T & vec_year %in% unique(na_yr)[j]),
                      which(colnames(x) %in% imp_med[i])]<-median(x[[which(colnames(x) %in% imp_med[i])]],na.rm = T)
                    cat("***All values in",imp_med[i],"for the year",unique(na_yr)[j],"are missing***\n")
                    cat("The median value across all years in the data has been used for imputation\n")
                  }
                  else {
                    x[which(is.na(x[[which(colnames(x) %in% imp_med[i])]]) == T & vec_year %in% unique(na_yr)[j]),
                      which(colnames(x) %in% imp_med[i])]<-median(x[[which(colnames(x) %in% imp_med[i])]][which(vec_year %in% unique(na_yr)[j])],na.rm = T)
                  }
                }
                log_na_mod<-x[[imp_med[i]]][na_row]
                dt_log_mis_med<-data.table("Variable" = rep(x = imp_med[i],times = length(na_row)),
                                           "Period" = vec_time[na_row],"Treatment" = rep(x = "Missing",times = length(na_row)),
                                           "Treatment technique" = rep(x = "Same year's median",times = length(na_row)),
                                           "Actual" = log_na_act,"Treated" = log_na_mod)
                dt_log_mis<-rbind.data.frame(dt_log_mis,dt_log_mis_med)
              }
            }
          }
          else if(imp_sel == "Impute with mean of same year"){
            imp_mean<-dlgList(choices = c(na_var,"---None---"),title = "Choose the variables to impute with mean",multiple = T)$res
            if("---None---" %in% imp_mean == F){
              imp_mean<-unlist(lapply(strsplit(x = imp_mean,split = " ~ ",fixed = T),function(x){x[[1]]}))
              vec_year<-unlist(lapply(strsplit(x = vec_time,split = " - ",fixed = T),function(x){x[[1]]}))
              for(i in 1:length(imp_mean)){
                na_row<-which(is.na(x[[which(colnames(x) %in% imp_mean[i])]]) == T)
                na_yr<-vec_year[na_row]
                log_na_act<-x[[imp_mean[i]]][na_row]
                for(j in 1:length(unique(na_yr))){
                  if(length(x[[which(colnames(x) %in% imp_mean[i])]][which(vec_year %in% unique(na_yr)[j])]) == sum(is.na(x[[which(colnames(x) %in% imp_mean[i])]][which(vec_year %in% unique(na_yr)[j])]))){
                    x[which(is.na(x[[which(colnames(x) %in% imp_mean[i])]]) == T & vec_year %in% unique(na_yr)[j]),
                      which(colnames(x) %in% imp_mean[i])]<-mean(x[[which(colnames(x) %in% imp_mean[i])]],na.rm = T)
                    cat("***All values in",imp_mean[i],"for the year",unique(na_yr)[j],"are missing***\n")
                    cat("The mean value across all years in the data has been used for imputation\n")
                  }
                  else {
                    x[which(is.na(x[[which(colnames(x) %in% imp_mean[i])]]) == T & vec_year %in% unique(na_yr)[j]),
                      which(colnames(x) %in% imp_mean[i])]<-mean(x[[which(colnames(x) %in% imp_mean[i])]][which(vec_year %in% unique(na_yr)[j])],na.rm = T)
                  }
                }
                log_na_mod<-x[[imp_mean[i]]][na_row]
                dt_log_mis_mean<-data.table("Variable" = rep(x = imp_mean[i],times = length(na_row)),
                                            "Period" = vec_time[na_row],"Treatment" = rep(x = "Missing",times = length(na_row)),
                                            "Treatment technique" = rep(x = "Same year's mean",times = length(na_row)),
                                            "Actual" = log_na_act,"Treated" = log_na_mod)
                dt_log_mis<-rbind.data.frame(dt_log_mis,dt_log_mis_mean)
              }
            }
          }
          else if(imp_sel == "Imputation with Rolling mean"){
            imp_roll<-dlgList(choices = c(na_var,"---None---"),title = "Choose the variables to impute with rolling mean",multiple = T)$res
            if("---None---" %in% imp_roll == F){
              roll_val<-as.numeric(dlgInput(message = "Enter the number of months/periods to be used for Rolling Mean calculation",
                                            default = 3)$res)
              imp_roll<-unlist(lapply(strsplit(x = imp_roll,split = " ~ ",fixed = T),function(x){x[[1]]}))
              for(i in 1:length(imp_roll)){
                roll_mn<-x[[which(colnames(x) %in% imp_roll[i])]]
                na_row<-which(is.na(roll_mn) == T)
                log_na_act<-x[[imp_roll[i]]][na_row]
                for(j in na_row){
                  if(j == 1){
                    ini_val<-NA
                    inc_fac<-1
                    while(is.na(ini_val) == T){
                      ini_val<-roll_mn[j + inc_fac]
                      inc_fac<-inc_fac + 1
                    }
                    roll_mn[j]<-ini_val
                  }
                  if(j != 1 & j < (roll_val + 1)){
                    pre_val<-roll_mn[j - 1]
                    suc_val<-NA
                    inc_fac<-1
                    while(is.na(suc_val) == T & (j + inc_fac) <= length(roll_mn)){
                      suc_val<-roll_mn[j + inc_fac]
                      inc_fac<-inc_fac + 1
                    }
                    if(is.na(suc_val) == T){
                      suc_val<-0
                    }
                    roll_mn[j]<-mean(x = c(pre_val,suc_val),na.rm = T)
                  }
                  else {
                    roll_mn[j]<-mean(roll_mn[(j - roll_val):(j - 1)],na.rm = T)
                  }
                }
                x[[which(colnames(x) %in% imp_roll[i])]]<-roll_mn
                log_na_mod<-x[[imp_roll[i]]][na_row]
                dt_log_mis_roll<-data.table("Variable" = rep(x = imp_roll[i],times = length(na_row)),
                                            "Period" = vec_time[na_row],"Treatment" = rep(x = "Missing",times = length(na_row)),
                                            "Treatment technique" = rep(x = "Rolling mean",times = length(na_row)),
                                            "Actual" = log_na_act,"Treated" = log_na_mod)
                dt_log_mis<-rbind.data.frame(dt_log_mis,dt_log_mis_roll)
              }
            }
          }
          else if(imp_sel == "Impute with fixed value"){
            imp_fix<-dlgList(choices = c(na_var,"---None---"),title = "Choose the variables to impute with fixed value",multiple = T)$res
            if("---None---" %in% imp_fix == F){
              imp_fix<-unlist(lapply(strsplit(x = imp_fix,split = " ~ ",fixed = T),function(x){x[[1]]}))
              val<-as.numeric(dlgInput(message = "Enter the value (numeric) to replace missing values",default = 0)$res)
              for(i in 1:length(imp_fix)){
                na_row<-which(is.na(x[[which(colnames(x) %in% imp_fix[i])]]) == T)
                log_na_act<-x[[imp_fix[i]]][na_row]
                x[na_row,which(colnames(x) %in% imp_fix[i])]<-val
                log_na_mod<-x[[imp_fix[i]]][na_row]
                dt_log_mis_fix<-data.table("Variable" = rep(x = imp_fix[i],times = length(na_row)),
                                           "Period" = vec_time[na_row],"Treatment" = rep(x = "Missing",times = length(na_row)),
                                           "Treatment technique" = rep(x = "Fixed value",times = length(na_row)),
                                           "Actual" = log_na_act,"Treated" = log_na_mod)
                dt_log_mis<-rbind.data.frame(dt_log_mis,dt_log_mis_fix)
              }
            }
          }
          if(sum(is.na(x)) > 0){
            imp<-dlgMessage(message = c(paste("No. of missing values remaining in the data:",sum(is.na(x)),sep = " "),
                                        "Do you want to continue imputing?"),type = "yesno")$res
          }
        }
        cat("Missing values were successfully imputed\n\n")

        return(list("Data" = x,"Log Miss" = dt_log_mis))
      }
      if(prm == T){
        for(i in 1:length(ls_dt_sub)){
          nm_ret<-names(ls_dt_sub)[i]
          for(j in 1:length(ls_dt_sub[[i]])){
            nm_ls<-paste(nm_ret,names(ls_dt_sub[[i]])[j],sep = "_")
            cat("Working on - ",nm_ls,"\n",sep = "")
            res_na_trt<-treat_NA(x = ls_dt_sub[[i]][[j]],name = nm_ls)
            ls_dt_sub[[i]][[j]]<-res_na_trt$Data
            dt_log_na<-cbind.data.frame(cbind.data.frame("Market/Retailer" = rep(x = nm_ret,times = nrow(res_na_trt$`Log Miss`)),
                                                         "Promo" = rep(x = names(ls_dt_sub[[i]])[j],times = nrow(res_na_trt$`Log Miss`))),
                                        res_na_trt$`Log Miss`)
            dt_log_trt<-rbind.data.frame(dt_log_trt,dt_log_na)
          }
        }
      }
      else {
        for(i in 1:length(ls_dt_sub)){
          nm_ls<-names(ls_dt_sub)[i]
          cat("Working on - ",nm_ls,"\n",sep = "")
          res_na_trt<-treat_NA(x = ls_dt_sub[[i]],name = nm_ls)
          ls_dt_sub[[i]]<-res_na_trt$Data
          dt_log_na<-cbind.data.frame("Market/Retailer" = rep(x = nm_ls,times = nrow(res_na_trt$`Log Miss`)),
                                      res_na_trt$`Log Miss`)
          dt_log_trt<-rbind.data.frame(dt_log_trt,dt_log_na)
        }
      }
      cat("-----Missing value imputation complete-----\n\n")
    }

    else {
      cat("-----No missing values are present in the data-----\n\n")
    }

    #####################################---Negative value treatment---#####################################

    treat_neg<-function(x,name){
      dt_neg<-x[,lapply(X = .SD,FUN = function(y){
        length(which(y < 0))
      }),.SDcols = colnames(x)]
      neg_col<-which(as.numeric(dt_neg[1,]) > 0)
      all_neg<-which(as.numeric(dt_neg[1,]) == nrow(x))
      dt_log_neg<-data.table("Variable" = character(),"Period" = character(),"Treatment" = character(),
                             "Treatment technique" = character(),"Actual" = numeric(),"Treated" = numeric())

      if(length(all_neg) > 0){
        dlgMessage(c("Variables with negative values throughout have been detected",
                     "Treat them to avoid errors in the module"),type = "okcancel")
      }

      if(sum(as.numeric(dt_neg[1,])) > 0){
        cat("Negative values are present in the data\nChoose the technique to replace them\n\n")
        rep_neg<-"yes"

        while(length(neg_col) > 0 & rep_neg == "yes"){
          dt_neg<-dt_neg[,neg_col,with = F]
          neg_var<-character(length = ncol(dt_neg))
          for(i in 1:ncol(dt_neg)){
            neg_pct<-round(dt_neg[[i]][1]/nrow(x) * 100,digits = 2)
            neg_var[i]<-paste(colnames(dt_neg)[i],paste(neg_pct,"% negative",sep = ""),paste(dt_neg[[i]][1],"values",sep = " "),sep = " ~ ")
          }
          rep_sel<-dlgList(choices = c("Replace with median of same year","Replace with mean of same year","Replace with Rolling mean","Replace with fixed value","Replace with absolute"),
                           title = "Choose replacement technique")$res

          if(rep_sel == "Replace with median of same year"){
            rep_med<-dlgList(choices = c(neg_var,"---None---"),title = "Choose variables to replace with median",multiple = T)$res
            if("---None---" %in% rep_med == F){
              rep_med<-unlist(lapply(strsplit(x = rep_med,split = " ~ ",fixed = T),function(x){x[[1]]}))
              vec_year<-unlist(lapply(strsplit(x = vec_time,split = " - ",fixed = T),function(x){x[[1]]}))
              for(i in 1:length(rep_med)){
                neg_row<-which(x[[which(colnames(x) %in% rep_med[i])]] < 0)
                neg_yr<-vec_year[neg_row]
                log_neg_act<-x[[rep_med[i]]][neg_row]
                for(j in 1:length(unique(neg_yr))){
                  if(length(x[[which(colnames(x) %in% rep_med[i])]][which(vec_year %in% unique(neg_yr)[j])]) == length(which(x[[which(colnames(x) %in% rep_med[i])]][which(vec_year %in% unique(neg_yr)[j])] < 0))){
                    x[which(x[[which(colnames(x) %in% rep_med[i])]] < 0 & vec_year %in% unique(neg_yr)[j]),
                      which(colnames(x) %in% rep_med[i])]<-median(x[[which(colnames(x) %in% rep_med[i])]][which(x[[which(colnames(x) %in% rep_med[i])]] >= 0)],na.rm = T)
                    cat("***All values in",rep_med[i],"for the year",unique(neg_yr)[j],"are negative***\n")
                    cat("The median value across all years in the data has been used for replacement\n")
                  }
                  else {
                    yr_sub<-x[[which(colnames(x) %in% rep_med[i])]][which(vec_year %in% unique(neg_yr)[j])]
                    x[which(x[[which(colnames(x) %in% rep_med[i])]] < 0 & vec_year %in% unique(neg_yr)[j]),
                      which(colnames(x) %in% rep_med[i])]<-median(yr_sub[which(yr_sub >= 0)],na.rm = T)
                  }
                }
                log_neg_mod<-x[[rep_med[i]]][neg_row]
                dt_log_neg_med<-data.table("Variable" = rep(x = rep_med[i],times = length(neg_row)),
                                           "Period" = vec_time[neg_row],"Treatment" = rep(x = "Negative",times = length(neg_row)),
                                           "Treatment technique" = rep(x = "Same year's median",times = length(neg_row)),
                                           "Actual" = log_neg_act,"Treated" = log_neg_mod)
                dt_log_neg<-rbind.data.frame(dt_log_neg,dt_log_neg_med)
              }
            }
          }
          else if(rep_sel == "Replace with mean of same year"){
            rep_mean<-dlgList(choices = c(neg_var,"---None---"),title = "Choose variables to replace with mean",multiple = T)$res
            if("---None---" %in% rep_mean == F){
              rep_mean<-unlist(lapply(strsplit(x = rep_mean,split = " ~ ",fixed = T),function(x){x[[1]]}))
              vec_year<-unlist(lapply(strsplit(x = vec_time,split = " - ",fixed = T),function(x){x[[1]]}))
              for(i in 1:length(rep_mean)){
                neg_row<-which(x[[which(colnames(x) %in% rep_mean[i])]] < 0)
                neg_yr<-vec_year[neg_row]
                log_neg_act<-x[[rep_mean[i]]][neg_row]
                for(j in 1:length(unique(neg_yr))){
                  if(length(x[[which(colnames(x) %in% rep_mean[i])]][which(vec_year %in% unique(neg_yr)[j])]) == length(which(x[[which(colnames(x) %in% rep_mean[i])]][which(vec_year %in% unique(neg_yr)[j])] < 0))){
                    x[which(x[[which(colnames(x) %in% rep_mean[i])]] < 0 & vec_year %in% unique(neg_yr)[j]),
                      which(colnames(x) %in% rep_mean[i])]<-mean(x[[which(colnames(x) %in% rep_mean[i])]][which(x[[which(colnames(x) %in% rep_med[i])]] >= 0)],na.rm = T)
                    cat("***All values in",rep_mean[i],"for the year",unique(neg_yr)[j],"are negative***\n")
                    cat("The mean value across all years in the data has been used for replacement\n")
                  }
                  else {
                    yr_sub<-x[[which(colnames(x) %in% rep_mean[i])]][which(vec_year %in% unique(neg_yr)[j])]
                    x[which(x[[which(colnames(x) %in% rep_mean[i])]] < 0 & vec_year %in% unique(neg_yr)[j]),
                      which(colnames(x) %in% rep_mean[i])]<-mean(yr_sub[which(yr_sub >= 0)],na.rm = T)
                  }
                }
                log_neg_mod<-x[[rep_mean[i]]][neg_row]
                dt_log_neg_mean<-data.table("Variable" = rep(x = rep_mean[i],times = length(neg_row)),
                                            "Period" = vec_time[neg_row],"Treatment" = rep(x = "Negative",times = length(neg_row)),
                                            "Treatment technique" = rep(x = "Same year's mean",times = length(neg_row)),
                                            "Actual" = log_neg_act,"Treated" = log_neg_mod)
                dt_log_neg<-rbind.data.frame(dt_log_neg,dt_log_neg_mean)
              }
            }
          }
          else if(rep_sel == "Replace with Rolling mean"){
            rep_roll<-dlgList(choices = c(neg_var,"---None---"),title = "Choose variables to replace with rolling mean",multiple = T)$res
            roll_val<-as.numeric(dlgInput(message = "Enter the number of months/periods to be used for Rolling Mean calculation",
                                          default = 3)$res)
            if("---None---" %in% rep_roll == F){
              rep_roll<-unlist(lapply(strsplit(x = rep_roll,split = " ~ ",fixed = T),function(x){x[[1]]}))
              for(i in 1:length(rep_roll)){
                roll_mn<-x[[which(colnames(x) %in% rep_roll[i])]]
                neg_row<-which(roll_mn < 0)
                log_neg_act<-roll_mn[neg_row]
                for(j in neg_row){
                  if(j == 1){
                    ini_val<-(-999)
                    inc_fac<-1
                    while(ini_val < 0){
                      ini_val<-roll_mn[j + inc_fac]
                      inc_fac<-inc_fac + 1
                    }
                    roll_mn[j]<-ini_val
                  }
                  if(j != 1 & j < (roll_val + 1)){
                    pre_val<-roll_mn[j - 1]
                    suc_val<-(-999)
                    inc_fac<-1
                    while(suc_val < 0 & (j + inc_fac) <= length(roll_mn)){
                      suc_val<-roll_mn[j + inc_fac]
                      inc_fac<-inc_fac + 1
                    }
                    if(suc_val < 0){
                      suc_val<-0
                    }
                    roll_mn[j]<-mean(x = c(pre_val,suc_val),na.rm = T)
                  }
                  else {
                    roll_mn[j]<-mean(roll_mn[(j - roll_val):(j - 1)],na.rm = T)
                  }
                }
                x[[which(colnames(x) %in% rep_roll[i])]]<-roll_mn
                log_neg_mod<-x[[rep_roll[i]]][neg_row]
                dt_log_neg_roll<-data.table("Variable" = rep(x = rep_roll[i],times = length(neg_row)),
                                            "Period" = vec_time[neg_row],"Treatment" = rep(x = "Negative",times = length(neg_row)),
                                            "Treatment technique" = rep(x = "Rolling mean",times = length(neg_row)),
                                            "Actual" = log_neg_act,"Treated" = log_neg_mod)
                dt_log_neg<-rbind.data.frame(dt_log_neg,dt_log_neg_roll)
              }
            }
          }
          else if(rep_sel == "Replace with fixed value"){
            rep_fix<-dlgList(choices = c(neg_var,"---None---"),title = "Choose variables to replace with fixed value",multiple = T)$res
            if("---None---" %in% rep_fix == F){
              rep_fix<-unlist(lapply(strsplit(x = rep_fix,split = " ~ ",fixed = T),function(x){x[[1]]}))
              val<-as.numeric(dlgInput(message = "Enter the value (numeric) to replace negative values",default = 0)$res)
              for(i in 1:length(rep_fix)){
                neg_row<-which(x[[which(colnames(x) %in% rep_fix[i])]] < 0)
                log_neg_act<-x[[rep_fix[i]]][neg_row]
                x[neg_row,which(colnames(x) %in% rep_fix[i])]<-val
                log_neg_mod<-x[[rep_fix[i]]][neg_row]
                dt_log_neg_fix<-data.table("Variable" = rep(x = rep_fix[i],times = length(neg_row)),
                                           "Period" = vec_time[neg_row],"Treatment" = rep(x = "Negative",times = length(neg_row)),
                                           "Treatment technique" = rep(x = "Fixed value",times = length(neg_row)),
                                           "Actual" = log_neg_act,"Treated" = log_neg_mod)
                dt_log_neg<-rbind.data.frame(dt_log_neg,dt_log_neg_fix)
              }
            }
          }
          else if(rep_sel == "Replace with absolute"){
            rep_abs<-dlgList(choices = c(neg_var,"---None---"),title = "Choose variables to replace with absolute value",multiple = T)$res
            if("---None---" %in% rep_abs == F){
              rep_abs<-unlist(lapply(strsplit(x = rep_abs,split = " ~ ",fixed = T),function(x){x[[1]]}))
              for(i in 1:length(rep_abs)){
                neg_row<-which(x[[rep_abs[i]]] < 0)
                log_neg_act<-x[[rep_abs[i]]][neg_row]
                x[[rep_abs[i]]]<-abs(x[[rep_abs[i]]])
                log_neg_mod<-x[[rep_abs[i]]][neg_row]
                dt_log_neg_abs<-data.table("Variable" = rep(x = rep_abs[i],times = length(neg_row)),
                                           "Period" = vec_time[neg_row],"Treatment" = rep(x = "Negative",times = length(neg_row)),
                                           "Treatment technique" = rep(x = "Absolute",times = length(neg_row)),
                                           "Actual" = log_neg_act,"Treated" = log_neg_mod)
                dt_log_neg<-rbind.data.frame(dt_log_neg,dt_log_neg_abs)
              }
            }
          }

          dt_neg<-x[,lapply(X = .SD,FUN = function(y){
            length(which(y < 0))
          }),.SDcols = colnames(x)]
          neg_col<-which(as.numeric(dt_neg[1,]) > 0)
          if(length(neg_col) > 0){
            rep_neg<-dlgMessage(message = c(paste("No. of negative values remaining in the data:",sum(as.numeric(dt_neg[1,])),sep = " "),
                                            "Do you want to continue replacing them?"),type = "yesno")$res
          }
        }
        cat("Negative values were successfully replaced\n\n")
      }

      else {
        cat("No negative values are present in the data\n\n")
      }

      return(list("Data" = x,"Log Neg" = dt_log_neg))
    }
    if(prm == T){
      for(i in 1:length(ls_dt_sub)){
        nm_ret<-names(ls_dt_sub)[i]
        for(j in 1:length(ls_dt_sub[[i]])){
          nm_ls<-paste(nm_ret,names(ls_dt_sub[[i]])[j],sep = "_")
          cat("Working on - ",nm_ls,"\n",sep = "")
          res_neg_trt<-treat_neg(x = ls_dt_sub[[i]][[j]],name = nm_ls)
          ls_dt_sub[[i]][[j]]<-res_neg_trt$Data
          dt_log_neg<-cbind.data.frame(cbind.data.frame("Market/Retailer" = rep(x = nm_ret,times = nrow(res_neg_trt$`Log Neg`)),
                                                        "Promo" = rep(x = names(ls_dt_sub[[i]])[j],times = nrow(res_neg_trt$`Log Neg`))),
                                       res_neg_trt$`Log Neg`)
          dt_log_trt<-rbind.data.frame(dt_log_trt,dt_log_neg)
        }
      }
    }
    else {
      for(i in 1:length(ls_dt_sub)){
        nm_ls<-names(ls_dt_sub)[i]
        cat("Working on - ",nm_ls,"\n",sep = "")
        res_neg_trt<-treat_neg(x = ls_dt_sub[[i]],name = nm_ls)
        ls_dt_sub[[i]]<-res_neg_trt$Data
        dt_log_neg<-cbind.data.frame("Market/Retailer" = rep(x = nm_ls,times = nrow(res_neg_trt$`Log Neg`)),
                                     res_neg_trt$`Log Neg`)
        dt_log_trt<-rbind.data.frame(dt_log_trt,dt_log_neg)
      }
    }
    cat("-----Negative value treatment complete-----\n\n")

    #####################################---Outlier analysis---#####################################

    treat_out<-function(x){
      ls_out<-lapply(X = x,FUN = function(y){
        qnt<-quantile(x = y,na.rm = T)
        low_lim<-qnt[2] - (1.5 * IQR(x = y,na.rm = T))
        up_lim<-qnt[4] + (1.5 * IQR(x = y,na.rm = T))
        outs<-which(y < low_lim | y > up_lim)

        return(outs)
      })
      ls_out<-ls_out[lapply(X = ls_out,FUN = length) > 0]

      dt_log_out<-data.table("Variable" = character(),"Period" = character(),"Treatment" = character(),
                             "Treatment technique" = character(),"Actual" = numeric(),"Treated" = numeric())

      if(length(unlist(ls_out)) > 0){
        chk_out<-dlgMessage(c("Outliers have been detected in the data","Follow the upcoming prompts to treat them"),
                            type = "okcancel")$res
        if(chk_out == "cancel"){
          stop("Function execution has been stopped")
        }
        out_var<-vector(length = length(ls_out))
        for(i in 1:length(ls_out)){
          out_var[i]<-paste(names(ls_out)[i],paste(length(ls_out[[i]])," outliers (",round(length(ls_out[[i]])/nrow(x) * 100,digits = 2),"%)",sep = ""),
                            paste("Values: ",paste(unique(round(x[[names(ls_out)[i]]][ls_out[[i]]],digits = 2)),collapse = ", "),sep = ""),
                            paste("Median: ",round(median(x[[names(ls_out)[i]]],na.rm = T),digits = 2),sep = ""),sep = " ~ ")
        }
        cat("Choose the variables to treat\n")
        out_imp<-dlgList(choices = c(out_var,"---None---"),multiple = T,title = "Choose variables to treat")$res
        if("---None---" %in% out_imp == F){
          out_imp<-sapply(X = strsplit(x = out_imp,split = " ~ ",fixed = T),FUN = function(x){x[[1]]})
          for(i in 1:length(out_imp)){
            outs<-ls_out[[which(names(ls_out) %in% out_imp[i])]]
            log_out_act<-x[[out_imp[i]]][outs]
            for(j in outs){
              if(j == 1){
                x[[out_imp[i]]][j]<-x[[out_imp[i]]][-outs][1]
              }
              else if(j != 1 & j != nrow(x)){
                obs_id<-1:nrow(x)
                nxt_obs<-obs_id[which(obs_id %in% outs == F & obs_id > j)][1]
                if(is.na(nxt_obs) == T){
                  nxt_obs<-(j - 1)
                }
                x[[out_imp[i]]][j]<-mean(x[[out_imp[i]]][j - 1],x[[out_imp[i]]][nxt_obs],na.rm = T)
              }
              else if(j == nrow(x)){
                x[[out_imp[i]]][j]<-x[[out_imp[i]]][j - 1]
              }
            }
            log_out_mod<-x[[out_imp[i]]][outs]
            dt_log_out_nb<-data.table("Variable" = rep(x = out_imp[i],times = length(outs)),
                                      "Period" = vec_time[outs],"Treatment" = rep(x = "Outlier",times = length(outs)),
                                      "Treatment technique" = rep(x = "Neighbor's mean",times = length(outs)),
                                      "Actual" = log_out_act,"Treated" = log_out_mod)
            dt_log_out<-rbind.data.frame(dt_log_out,dt_log_out_nb)
          }
          cat("Outliers have been successfully treated\n\n")
        }
        else {
          cat("Outliers have been retained in the data\n\n")
        }
      }

      else {
        cat("No outliers are present in the data\n\n")
      }

      return(list("Data" = x,"Log Out" = dt_log_out))
    }

    if(prm == T){
      for(i in 1:length(ls_dt_sub)){
        nm_ret<-names(ls_dt_sub)[i]
        for(j in 1:length(ls_dt_sub[[i]])){
          nm_ls<-paste(nm_ret,names(ls_dt_sub[[i]])[j],sep = " : ")
          cat("Working on - ",nm_ls,"\n",sep = "")
          res_out_trt<-treat_out(x = ls_dt_sub[[i]][[j]])
          ls_dt_sub[[i]][[j]]<-res_out_trt$Data
          dt_log_out<-cbind.data.frame(cbind.data.frame("Market/Retailer" = rep(x = nm_ret,times = nrow(res_out_trt$`Log Out`)),
                                                        "Promo" = rep(x = names(ls_dt_sub[[i]])[j],times = nrow(res_out_trt$`Log Out`))),
                                       res_out_trt$`Log Out`)
          dt_log_trt<-rbind.data.frame(dt_log_trt,dt_log_out)
          dt_log_trt<-dt_log_trt[order(`Market/Retailer`,Promo,Variable,decreasing = F)]
        }
      }
    }
    else {
      for(i in 1:length(ls_dt_sub)){
        nm_ls<-names(ls_dt_sub)[i]
        cat("Working on - ",nm_ls,"\n",sep = "")
        res_out_trt<-treat_out(x = ls_dt_sub[[i]])
        ls_dt_sub[[i]]<-res_out_trt$Data
        dt_log_out<-cbind.data.frame("Market/Retailer" = rep(x = nm_ls,times = nrow(res_out_trt$`Log Out`)),
                                     res_out_trt$`Log Out`)
        dt_log_trt<-rbind.data.frame(dt_log_trt,dt_log_out)
        dt_log_trt<-dt_log_trt[order(`Market/Retailer`,Variable,decreasing = F)]
      }
    }

    cat("-----Outlier treatment successfully completed-----\n\n")

    #####################################---Constant treatment---#####################################

    treat_con<-function(x,name){
      dt_con<-x[,lapply(X = .SD,FUN = function(x){
        length(unique(x))
      }),.SDcols = colnames(x)]
      con_col<-colnames(x)[which(as.numeric(dt_con[1,]) == 1)]
      if(length(con_col) > 0){
        dlgMessage(c("Variables with constant value throughout have been detected",
                     "Details will be saved in your working directory"),type = "okcancel")
        writeLines(text = c("Variables with only constant value:",con_col),
                   con = paste(wd_path,paste("Var_Con_",name,".txt",sep = ""),sep = "\\"))
      }
      else {
        cat("No constant variables present\n\n")
      }
      return(x)
    }

    if(prm == T){
      for(i in 1:length(ls_dt_sub)){
        nm_ret<-names(ls_dt_sub)[i]
        for(j in 1:length(ls_dt_sub[[i]])){
          nm_ls<-paste(nm_ret,names(ls_dt_sub[[i]])[j],sep = "_")
          cat("Working on - ",nm_ls,"\n",sep = "")
          ls_dt_sub[[i]][[j]]<-treat_con(x = ls_dt_sub[[i]][[j]],name = nm_ls)
        }
      }
    }
    else {
      for(i in 1:length(ls_dt_sub)){
        nm_ls<-names(ls_dt_sub)[i]
        cat("Working on - ",nm_ls,"\n",sep = "")
        ls_dt_sub[[i]]<-treat_con(x = ls_dt_sub[[i]],name = nm_ls)
      }
    }

    cat("-----Constant variable treatment successfully completed-----\n\n")

    #####################################---Output generation---#####################################

    cat("-----Generating outputs-----\n\n")

    make_out<-function(x,prm_fam,ret_nm){
      if(exists("vec_time")){
        x<-cbind.data.frame(cbind.data.frame("Year" = sapply(X = strsplit(x = vec_time,split = " - ",fixed = T),FUN = function(y){
          as.numeric(y[[1]])
        }),"Period" = as.numeric(sapply(X = strsplit(x = vec_time,split = " - ",fixed = T),FUN = function(y){
          per_num<-sapply(X = strsplit(x = y[[2]],split = ":",fixed = T),FUN = function(z){
            z[[2]]
          })
          return(per_num)
        }))),x)
      }
      else {
        x<-cbind.data.frame(cbind.data.frame("Year" = rep(x = NA,times = nrow(x)),
                                             "Period" = rep(x = NA,times = nrow(x))),x)
      }
      if(prm == T){
        x<-cbind.data.frame("Promo family" = rep(x = prm_fam,times = nrow(x)),x)
      }
      else {
        x<-cbind.data.frame("Promo family" = rep(x = NA,times = nrow(x)),x)
      }
      x<-cbind.data.frame("Market/Retailer" = rep(x = ret_nm,times = nrow(x)),x)

      return(x)
    }

    ls_out_trt<-vector(mode = "list",length(ls_dt_sub))
    names(ls_out_trt)<-names(ls_dt_sub)

    if(prm == T){
      for(i in 1:length(ls_dt_sub)){
        ls_prm<-vector(length = length(ls_dt_sub[[i]]),mode = "list")
        names(ls_prm)<-names(ls_dt_sub[[i]])
        nm_ret<-names(ls_dt_sub)[i]
        for(j in 1:length(ls_dt_sub[[i]])){
          nm_prm<-names(ls_dt_sub[[i]])[j]
          ls_prm[[j]]<-make_out(x = ls_dt_sub[[i]][[j]],ret_nm = nm_ret,prm_fam = nm_prm)
        }
        ls_out_trt[[i]]<-ls_prm
      }
      ls_out_trt<<-lapply(X = ls_out_trt,FUN = rbindlist,use.names = T)
      dt_out_trt<-rbindlist(l = ls_out_trt,use.names = T)
      dt_out_trt<-dt_out_trt[order(`Market/Retailer`,`Promo family`,Year,Period,decreasing = F)]
    }
    else {
      for(i in 1:length(ls_dt_sub)){
        nm_ls<-names(ls_dt_sub)[i]
        ls_out_trt[[i]]<-make_out(x = ls_dt_sub[[i]],ret_nm = nm_ls,prm_fam = NA)
      }
      dt_out_trt<-rbindlist(l = ls_out_trt,use.names = T)
      dt_out_trt<-dt_out_trt[order(`Market/Retailer`,`Promo family`,Year,Period,decreasing = F)]
    }

    fl_nm_ls<-unlist(strsplit(x = fl_path,split = "\\",fixed = T))
    wb_rep<-openxlsx::createWorkbook()
    openxlsx::addWorksheet(wb = wb_rep,sheetName = "Processed data",gridLines = T)
    openxlsx::writeDataTable(wb = wb_rep,sheet = "Processed data",x = dt_out_trt,colNames = T,rowNames = F)
    openxlsx::setColWidths(wb = wb_rep,sheet = "Processed data",cols = 1:ncol(dt_out_trt),widths = "auto")
    openxlsx::addWorksheet(wb = wb_rep,sheetName = "Treatment info",gridLines = T)
    openxlsx::writeDataTable(wb = wb_rep,sheet = "Treatment info",x = dt_log_trt,colNames = T,rowNames = F)
    openxlsx::setColWidths(wb = wb_rep,sheet = "Treatment info",cols = 1:ncol(dt_log_trt),widths = "auto")
    openxlsx::saveWorkbook(wb = wb_rep,file = paste(wd_path,paste(paste("Processed data",unlist(strsplit(x = fl_nm_ls[length(fl_nm_ls)],split = "\\."))[1],sep = "_")
                                                                  ,".xlsx",sep = ""),sep = "\\"),overwrite = T)

    dlgMessage(message = c("***Data processing complete***",paste("Your output files are saved in :",wd_path,sep = " ")),type = "okcancel")

    return(list("Data" = ls_dt_sub,"Period" = vec_time,"Scope" = dt_scp,"Working Directory" = wd_path))
  })
}

#' Automated data transformation and variable selection
#'
#' The function reads the processed data from Prepare_Data, creates transformed versions of selected variables,
#' selects the transformed version with the ideal correlation to the target and writes out the final train and
#' test splits of the data which can be directly fed to the Bayesian network creation module or any other models.
#' @param prm Logical. If set to TRUE, the functions filters the data for an additional category. For the US market,
#' use prm = TRUE to filter for both Retailer and Promo Family. For the India market, can be used to filter for both
#' Brand and Cluster in case consolidated data for multiple brands is provided in a single file.
#' @param ls_dt_sub The "Data" object returned by Prepare_Data
#' @param vec_time The "Period" object returned by Prepare_Data
#' @param dt_scp The "Scope" object returned by Prepare_Data
#' @param wd_path The "Working Directory" object returned by Prepare_Data
#' @return All outputs are generated within specific folders for the different Markets/Retailers and Promos
#' @import data.table svDialogs
#' @export

Transform_Data<-function(ls_dt_sub,vec_time,dt_scp,wd_path,prm = F){
  suppressMessages({
    suppressWarnings({
      require(svDialogs,quietly = T,warn.conflicts = F)
    })

    #####################################---Directory creation---#####################################

    for(i in 1:length(ls_dt_sub)){
      if(prm == T){
        for(j in 1:length(ls_dt_sub[[i]])){
          dir.create(path = paste(wd_path,paste(names(ls_dt_sub)[i],names(ls_dt_sub[[i]])[j],sep = "_"),sep = "\\"))
        }
      }
      else {
        dir.create(path = paste(wd_path,names(ls_dt_sub)[i],sep = "\\"))
      }
    }
    #####################################---Target selection---#####################################

    tar_var<-dt_scp$Scope[["Model_Name"]][which(dt_scp$Scope[["Variable_Type"]] == "Target")]
    cat("-----",tar_var," has been chosen as the target variable-----\n\n",sep = "")

    #####################################---Variable transformation---#####################################

    adstock<-function(dt,col,adstk){
      col_wk<-dt[[col]]
      col_ad<-numeric(length = length(col_wk))
      for(i in 1:length(col_wk)){
        if(i == 1){
          col_ad[i]<-col_wk[i]
        }
        else {
          col_ad[i]<-col_wk[i] + (col_ad[i - 1]*adstk)
        }
      }
      res_list<-list(col_ad)
      names(res_list)<-paste(col,paste("Ads",adstk*100,sep = ""),sep = ".")

      return(res_list)
    }

    lagFun<-function(dt,col,lagVal){
      col_wk<-dt[[col]]
      col_lag<-Hmisc::Lag(x = col_wk,shift = lagVal)
      col_lag[which(is.na(col_lag) == T)]<-col_wk[which(is.na(col_lag) == T)]
      res_list<-list(col_lag)
      names(res_list)<-paste(col,paste("lag",lagVal,sep = ""),sep = ".")

      return(res_list)
    }

    rollPrice<-function(dt,col,rollval){
      mma_col<-dt[[col]]
      res_list<-list(zoo::na.locf(object = dplyr::lag(x = zoo::rollapply(data = mma_col,width = rollval,
                                                                         align = "right",fill = NA,partial = T,
                                                                         FUN = mean),n = 1),fromLast = T))
      names(res_list)<-paste(paste(paste(col,"MMA",sep = "."),rollval,sep = ""),"lag1",sep = ".")
      return(res_list)
    }

    make_trfl<-function(x,trans = NULL){
      x<-x[is.na(eval(as.name(trans))) == F]
      dt_trfl<-data.table("Variable" = character(),"Value" = numeric())
      for(i in 1:nrow(x)){
        min<-as.numeric(unlist(strsplit(x = x[[trans]][i],split = ",",fixed = T))[1])
        max<-as.numeric(unlist(strsplit(x = x[[trans]][i],split = ",",fixed = T))[2])
        val_tr<-seq(from = min,to = max,by = 1)
        var_tr<-rep(x = x[["Model_Name"]][i],times = length(val_tr))
        dt_trfl<-rbind.data.frame(dt_trfl,data.table("Variable" = var_tr,"Value" = val_tr))
      }
      return(dt_trfl)
    }

    dt_lag<-make_trfl(x = dt_scp$Scope,trans = "Lag")
    dt_ads<-make_trfl(x = dt_scp$Scope,trans = "ADStock")
    dt_mma<-make_trfl(x = dt_scp$Scope,trans = "MMA")
    dt_sign<-dt_scp$Sign

    transform<-function(x){
      dt_lag<-dt_lag[Variable %in% colnames(x)]
      dt_ads<-dt_ads[Variable %in% colnames(x)]
      dt_mma<-dt_mma[Variable %in% colnames(x)]

      trl_list<-list()
      if(nrow(dt_lag) > 0){
        for(i in 1:nrow(dt_lag)){
          lag_res<-lagFun(dt = x,col = dt_lag[["Variable"]][i],lagVal = dt_lag[["Value"]][i])
          trl_list[i]<-lag_res
          names(trl_list)[i]<-names(lag_res)
        }
      }
      tra_list<-list()
      if(nrow(dt_ads) > 0){
        for(i in 1:nrow(dt_ads)){
          ads_res<-adstock(dt = x,col = dt_ads[["Variable"]][i],adstk = dt_ads[["Value"]][i])
          tra_list[i]<-ads_res
          names(tra_list)[i]<-names(ads_res)
        }
      }
      trm_list<-list()
      if(nrow(dt_mma) > 0){
        for(i in 1:nrow(dt_mma)){
          mma_res<-rollPrice(dt = x,col = dt_mma[["Variable"]][i],rollval = dt_mma[["Value"]][i])
          trm_list[i]<-mma_res
          names(trm_list)[i]<-names(mma_res)
        }
      }

      dt_all<-x
      if(nrow(dt_lag) > 0){
        dt_all<-cbind.data.frame(dt_all,as.data.table(trl_list))
        trl_nm<-sapply(X = strsplit(names(trl_list),split = ".lag",fixed = T),FUN = function(x){x[1]})
      }
      else {
        trl_nm<-character()
      }
      if(nrow(dt_ads) > 0){
        dt_all<-cbind.data.frame(dt_all,as.data.table(tra_list))
        tra_nm<-sapply(X = strsplit(names(tra_list),split = ".Ads",fixed = T),FUN = function(x){x[1]})
      }
      else {
        tra_nm<-character()
      }
      if(nrow(dt_mma) > 0){
        dt_all<-cbind.data.frame(dt_all,as.data.table(trm_list))
        trm_nm<-sapply(X = strsplit(names(trm_list),split = ".MMA",fixed = T),FUN = function(x){x[1]})
      }
      else {
        trm_nm<-character()
      }

      cor_res<-vector(mode = "list",length = ncol(x))
      sub_nm<-names(cor_res)<-colnames(x)
      cor_vec<-character()
      for(i in 1:length(sub_nm)){
        if(sub_nm[i] %in% trl_nm == T | sub_nm[i] %in% tra_nm == T | sub_nm[i] %in% trm_nm){
          chk_cor<-c(sub_nm[i],names(trl_list)[which(trl_nm %in% sub_nm[i] == T)],names(tra_list)[which(tra_nm %in% sub_nm[i] == T)],
                     names(trm_list)[which(trm_nm %in% sub_nm[i] == T)])
          cor_col<-numeric()
          for(j in 1:length(chk_cor)){
            suppressWarnings({
              cor_col<-c(cor_col,round(cor(x = dt_all[[chk_cor[j]]],y = dt_all[[tar_var]]),digits = 2))
              names(cor_col)[j]<-chk_cor[j]
            })
          }
          cor_res[[i]]<-cor_col
          if(sum(is.na(cor_col)) < length(cor_col)){
            exp_sign<-as.numeric(dt_scp$Sign[["Impact"]][which(dt_scp$Sign[["Variable"]] %in% sub_nm[i] == T)])
            if(exp_sign < 0){
              x[[i]]<-dt_all[[chk_cor[which.min(cor_col)]]]
              colnames(x)[i]<-chk_cor[which.min(cor_col)]
            }
            else {
              x[[i]]<-dt_all[[chk_cor[which.max(cor_col)]]]
              colnames(x)[i]<-chk_cor[which.max(cor_col)]
            }
          }
        }
      }

      return(list("Selected" = x,"All" = dt_all))
    }

    ls_dt_all<-ls_dt_sub
    if(prm == T){
      for(i in 1:length(ls_dt_sub)){
        nm_ret<-names(ls_dt_sub)[i]
        for(j in 1:length(ls_dt_sub[[i]])){
          nm_ls<-paste(nm_ret,names(ls_dt_sub[[i]])[j],sep = " : ")
          cat("Working on - ",nm_ls,"\n",sep = "")
          tr_res<-transform(x = ls_dt_sub[[i]][[j]])
          ls_dt_sub[[i]][[j]]<-tr_res$Selected
          ls_dt_all[[i]][[j]]<-tr_res$All
          cat("Transformation and variable selection complete\n\n")
        }
      }
    }
    else {
      for(i in 1:length(ls_dt_sub)){
        nm_ls<-names(ls_dt_sub)[i]
        cat("Working on - ",nm_ls,"\n",sep = "")
        tr_res<-transform(x = ls_dt_sub[[i]])
        ls_dt_sub[[i]]<-tr_res$Selected
        ls_dt_all[[i]]<-tr_res$All
        cat("Transformation and variable selection complete\n\n")
      }
    }

    cat("-----Data transformation and variable selection complete-----\n\n")

    #####################################---Plotting---#####################################

    plot_cor<-function(x,name,x_all){
      suppressWarnings(expr = {
        cor_mat<-cor(x = x)
        cor_mat_all<-cor(x = x_all)
        png(filename = paste(wd_path,paste(name,"\\Cor_plot_",name,".png",sep = ""),sep = "\\"),width = 1920,height = 1920)
        corrplot::corrplot(corr = cor_mat)
        dev.off()
      })
      colnames(cor_mat_all)<-colnames(x_all)
      cor_df<-cbind.data.frame("Variable" = colnames(x_all),cor_mat_all)
      openxlsx::write.xlsx(x = cor_df,file = paste(wd_path,paste(name,"\\Cor_mat_",name,".xlsx",sep = ""),sep = "\\"),asTable = F)
      return(x)
    }

    if(prm == T){
      for(i in 1:length(ls_dt_sub)){
        nm_ret<-names(ls_dt_sub)[i]
        for(j in 1:length(ls_dt_sub[[i]])){
          nm_ls<-paste(nm_ret,names(ls_dt_sub[[i]])[j],sep = "_")
          ls_dt_sub[[i]][[j]]<-plot_cor(x = ls_dt_sub[[i]][[j]],name = nm_ls,x_all = ls_dt_all[[i]][[j]])
        }
      }
    }
    else {
      for(i in 1:length(ls_dt_sub)){
        nm_ls<-names(ls_dt_sub)[i]
        ls_dt_sub[[i]]<-plot_cor(x = ls_dt_sub[[i]],name = nm_ls,x_all = ls_dt_all[[i]])
      }
    }

    cat("-----Correlation plots generated-----\n\n")

    #####################################---Log transformation---#####################################

    ls_dt_sub_raw<-ls_dt_sub
    log_transform<-function(y){
      y<-y[,lapply(X = .SD,FUN = function(x){
        x[which(x <= 0)]<-0.00001
        return(x)
      }),.SDcols = colnames(y)]
      y<-log(y)
      return(y)
    }
    if(prm == T){
      ls_dt_sub<-lapply(X = ls_dt_sub,FUN = function(y){
        lapply(X = y,FUN = log_transform)
      })
    }
    else {
      ls_dt_sub<-lapply(X = ls_dt_sub,FUN = log_transform)
    }

    #####################################---Train/Test split---#####################################

    pt_tst<-dlgList(choices = vec_time,title = "Select test data points",multiple = T)$res
    eda_var<-dt_scp$Scope[["Model_Name"]][which(dt_scp$Scope[["Variable_Type"]] %in% "EDA") == T]
    data_split<-function(x,raw = F,name){
      dt_tst<-x[which(vec_time %in% pt_tst),which(colnames(x) %in% eda_var == F)]
      dt_trn<-x[which(vec_time %in% pt_tst == F),which(colnames(x) %in% eda_var == F)]
      if(raw == F){
        write.csv(x = dt_trn,row.names = F,
                  file = paste(wd_path,paste(name,"\\Train_Log_",name,".csv",sep = ""),sep = "\\"))
        write.csv(x = dt_tst,row.names = F,
                  file = paste(wd_path,paste(name,"\\Test_Log_",name,".csv",sep = ""),sep = "\\"))
        writeLines(text = c("Testing data points:",pt_tst),
                   con = paste(wd_path,paste(name,"\\Test_Points_",name,".txt",sep = ""),sep = "\\"))
        writeLines(text = c("Training data points:",vec_time[which(vec_time %in% pt_tst == F)]),
                   con = paste(wd_path,paste(name,"\\Train_Points_",name,".txt",sep = ""),sep = "\\"))
      }
      else {
        write.csv(x = dt_trn,row.names = F,
                  file = paste(wd_path,paste(name,"\\Train_Raw_",name,".csv",sep = ""),sep = "\\"))
        write.csv(x = dt_tst,row.names = F,
                  file = paste(wd_path,paste(name,"\\Test_Raw_",name,".csv",sep = ""),sep = "\\"))
      }
      return(x)
    }
    if(prm == T){
      for(i in 1:length(ls_dt_sub)){
        nm_ret<-names(ls_dt_sub)[i]
        for(j in 1:length(ls_dt_sub[[i]])){
          nm_ls<-paste(nm_ret,names(ls_dt_sub[[i]])[j],sep = "_")
          ls_dt_sub[[i]][[j]]<-data_split(x = ls_dt_sub[[i]][[j]],name = nm_ls)
          ls_dt_sub_raw[[i]][[j]]<-data_split(x = ls_dt_sub_raw[[i]][[j]],name = nm_ls,raw = T)
        }
      }
    }
    else {
      for(i in 1:length(ls_dt_sub)){
        nm_ls<-names(ls_dt_sub)[i]
        ls_dt_sub[[i]]<-data_split(x = ls_dt_sub[[i]],name = nm_ls)
        ls_dt_sub_raw[[i]]<-data_split(x = ls_dt_sub_raw[[i]],name = nm_ls,raw = T)
      }
    }

    cat("-----Train and Test splits generated-----\n\n")

    #####################################---Output generation---#####################################

    write_out<-function(x,prm_fam,ret_nm,name){
      if(exists("vec_time")){
        x<-cbind.data.frame(cbind.data.frame("Year" = sapply(X = strsplit(x = vec_time,split = " - ",fixed = T),FUN = function(y){
          as.numeric(y[[1]])
        }),"Period" = as.numeric(sapply(X = strsplit(x = vec_time,split = " - ",fixed = T),FUN = function(y){
          per_num<-sapply(X = strsplit(x = y[[2]],split = ":",fixed = T),FUN = function(z){
            z[[2]]
          })
          return(per_num)
        }))),x)
      }
      else {
        x<-cbind.data.frame(cbind.data.frame("Year" = rep(x = NA,times = nrow(x)),
                                             "Period" = rep(x = NA,times = nrow(x))),x)
      }
      if(prm == T){
        x<-cbind.data.frame("Promo family" = rep(x = prm_fam,times = nrow(x)),x)
      }
      else {
        x<-cbind.data.frame("Promo family" = rep(x = NA,times = nrow(x)),x)
      }
      x<-cbind.data.frame("Market/Retailer" = rep(x = ret_nm,times = nrow(x)),x)

      write.xlsx(x = x,file = paste(wd_path,paste(name,"\\Transformed Data_",name,".xlsx",sep = ""),sep = "\\"))

      return(x)
    }

    if(prm == T){
      for(i in 1:length(ls_dt_sub)){
        nm_ret<-names(ls_dt_sub)[i]
        for(j in 1:length(ls_dt_sub[[i]])){
          nm_prm<-names(ls_dt_sub[[i]])[j]
          nm_ls<-paste(nm_ret,names(ls_dt_sub[[i]])[j],sep = "_")
          ls_dt_sub_raw[[i]][[j]]<-write_out(x = ls_dt_sub_raw[[i]][[j]],name = nm_ls,ret_nm = nm_ret,prm_fam = nm_prm)
        }
      }
    }
    else {
      for(i in 1:length(ls_dt_sub)){
        nm_ls<-names(ls_dt_sub)[i]
        ls_dt_sub_raw[[i]]<-write_out(x = ls_dt_sub_raw[[i]],name = nm_ls,ret_nm = nm_ls,prm_fam = NA)
      }
    }

    dlgMessage(message = c("***EDA complete***",paste("Your output files are saved in :",wd_path,sep = " ")),type = "okcancel")

  })
}
