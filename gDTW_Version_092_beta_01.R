 # install.packages("matlib")
 # install.packages("tidyverse")
 # install.packages("reprex") # Code를 Notion에 가져가기 위한 팩키지
 # install.packages("readxl")
 # install.packages("writexl")
 # install.packages("svDialogs")
 # install.packages("dtw") #완료 후 제거할 것.
 # install.packages("openxlsx")
 # install.packages("plyr")

# 내적 공식을 이용하기 위해서 matlib 라이브러리의 angle 함수를 사용하였으나, 일부 문제가 있어 내적 함수를 만들어 사용함. 
library(plyr)
library(tidyverse)
library(tools)
library(reprex)
library(readxl)
library(writexl)
library(svDialogs)
library(openxlsx)
library(matlib)
library(dtw) #완료 후 제거할 것.



#===엑셀파일의 첫번째 쉬트 자료를 읽어들여 data.frame으로 만든 다음, 각 열을 index1, index2.... 형태로 이름을 부여한다.
#===입력은 없으며, data.frame 형태로 리턴한다.
ft_ReadExcel = function(non_data_df, filetype = 0){
   
   if(filetype == 0){
      path = choose.files(caption = "엑셀 데이터 파일(.XLSX)을 선택하세요.")
   } else{
      path = choose.files(caption = "구간 분리 값이 있는 엑셀파일(.XLSX)을 선택하세요.")
   }
   
   sheet.name=excel_sheets(path) # 엑셀 파일 내의 쉬트 이름을 저장한다.
   length(sheet.name) # Counting the sheet number.
   
   #엑셀 파일의 첫 번째 쉬트에서 데이터를 읽어 df_data에 저장한다.
   df_data = data.frame(read_excel(path, sheet = sheet.name[1], col_names = TRUE, na = ""))
   index_name = c("index1", "index2")
   
   #읽어 들인 데이터의 모든 컬럼에 index1, index2.... 형태로 변수명을 입력하기 위해 index_name 배열을 생성한다.
   for(i in 1:ncol(df_data)){
      index_name[i]=c(paste0("index", i))
   }
   
   #df_data에 위에서 생성한 문자열 배열을 이용하여 변수명을 부여한다.
   names(df_data)=c(index_name)
   
   # 첫번째 쉬트의 데이터를 읽어들인 후 새탭에 보여준다. 
   #View(df_data)
   return(df_data)
}


#===local cost matrix를 만든다. 함수 내에 NA 값을 제거하는 코드가 들어가 있음.
ft_CostMatrix <- function(df_data_mt){ # data.frame을 받아서 matrix인 mt.cost_matrix를 만드는 함수
   
   stopifnot(is.data.frame(df_data_mt))
   data = df_data_mt
   names(data) = c("Reference", "Query")
   
   # NA 값을 제거하기 위해서 임시 저장한다.
   temp.ref = data$Reference
   temp.query = data$Query
   
   
   # NA 값을 제거한다.
   ref.data = temp.ref[!is.na(temp.ref)]
   query.data = temp.query[!is.na(temp.query)]
   
   # referenc와 query의 길이에 맞는 행렬을 미리 만들어 놓는다.
   mt_CostMat = matrix(nrow = length(ref.data), ncol = length(query.data))
   
   # Y축의 첫 번째 열 값을 구한다.
   Yaxis = length(ref.data)
   Xaxis = length(query.data)
   mt_CostMat[1, 1] = sqrt((query.data[1] - ref.data[1])^2)
   
   for (i in 2:Yaxis) {
      mt_CostMat[i,1] = sqrt((query.data[1] - ref.data[i])^2) + mt_CostMat[i-1, 1]
   }
   
   # X축의 첫 번째 행 값을 구한다.
   for (i in 2:Xaxis) {
      mt_CostMat[1, i] = sqrt((query.data[i] - ref.data[1])^2) + mt_CostMat[1, i-1]
   }
   
   # 나머지 셀들의 값을 구한다.
   for (i in 2:Yaxis) {
      for (j in 2:Xaxis) {
         mt_CostMat[i, j] = sqrt((ref.data[i] - query.data[j])^2) + min(mt_CostMat[i-1, j],
                                                                        mt_CostMat[i, j-1],mt_CostMat[i-1, j-1])
      }
   }
   
   return(mt_CostMat)
}


#===Find the warping path. Matrix를 받아서 첫 번째 warping path를 찾는다.
#===매트릭스를 인수로 받아 list로 리턴한다.
ft_FirstWarpingPath <- function(mt_data_lst, phaseRef=0, phaseQuery=0, frame.x=0, frame.y=0 ){

   #stopifnot(is.matrix(data))
   #stopifnot(is.integer(phaseRef))
   #stopifnot(is.integer(phaseQuery))
   
   cost.matrix = mt_data_lst
   
   #df_Index와 df_SameIndex를 정의한다.
   df_Index = data.frame(Line=numeric(), Phase=character(), Index=numeric(), RowIndex=numeric(), ColIndex=numeric(), Cost=numeric())
   df_SameIndex = data.frame(LineFrom=numeric(), LineTo=numeric(), OrigIndex=numeric(), S_Count=numeric(), S_RowIndex=numeric(), 
                             S_ColIndex=numeric(), S_Cost=numeric())
   
  
   # i행 j열로 처리함.
   i = length(cost.matrix[,1]) #y axis and row index
   j = length(cost.matrix[1,]) #x axis and column index
   
   
   # 가장 끝에 있는 값을 저장한다.
   linefrom_no = 1; subphase = "ALL"; lineto_no = 1; index_no = 1; same_no = 0; orig_no = 0
   
   df_temp = data.frame(Line=linefrom_no, Phase=subphase, Index=index_no, RowIndex=i, ColIndex=j, Cost=cost.matrix[i, j])
   df_Index = rbind(df_Index, df_temp)
   
   # 비교할 데이터를 만드는 sub 함수이다.
   ft_CompData <- function(x, y){
      
      i = x; j = y
      
      if(i > 1 & j > 1){ #X축과 Y축 값이 모두 1 이상인 경우. 즉 왼쪽과 아래쪽에 모두 데이터가 있는 경우.
         CompData = as.vector(c(cost.matrix[i, j-1], cost.matrix[i-1, j-1], cost.matrix[i-1, j])) # 반시계 방향으로 처리함.
      } else if(i == 1 & j > 1){
         CompData = c(cost.matrix[i, j-1], cost.matrix[i, j-1]+1, cost.matrix[i, j-1]+1) #왼쪽값만 입력하고 나머지는 1이 더 큰 값을 입력한다.
      } else if(i > 1 & j == 1){
         CompData = c(cost.matrix[i-1, j]+1, cost.matrix[i-1, j]+1, cost.matrix[i-1, j]) #아래값만 입력하고 나머지는 1이 큰 값을 입력한다.
      } else if(i == 1 & j == 1){ # 없어도 돼지만 귀찮아서 그냥 둠.
         break
      }
      return(CompData)
   }
   
   
   while (i >= 1 | j >= 1) { #맨 위에서부터 처리함.
      
      # 비교할 데이터를 Comp.data에 저장한다.
      # if(i > 1 & j > 1){ #X축과 Y축 값이 모두 1 이상인 경우. 즉 왼쪽과 아래쪽에 모두 데이터가 있는 경우.
      #    Comp.data = as.vector(c(cost.matrix[i, j-1], cost.matrix[i-1, j-1], cost.matrix[i-1, j])) # 반시계 방향으로 처리함.
      # } else if(i == 1 & j > 1){
      #    Comp.data = c(cost.matrix[i, j-1], cost.matrix[i, j-1]+1, cost.matrix[i, j-1]+1) #왼쪽값만 입력하고 나머지는 1이 더 큰 값을 입력한다.
      # } else if(i > 1 & j == 1){
      #    Comp.data = c(cost.matrix[i-1, j]+1, cost.matrix[i-1, j]+1, cost.matrix[i-1, j]) #아래값만 입력하고 나머지는 1이 큰 값을 입력한다.
      # } else if(i == 1 & j == 1){
      #    # index_no = index_no + 1
      #    # df_Index[index_no,] = c(linefrom_no, index_no, i, j, cost.matrix[i, j])
      #    # print(i); print(j)
      #    break
      # }
      
      if(i == 1 & j == 1){
         break
      }
      
      Comp.data = ft_CompData(i, j)
      
      if(length(which(Comp.data == Comp.data[which.min(Comp.data)])) == 1){ # 최소값이 중복되지 않은 경우.
         
         if (which.min(Comp.data) == 1){ # 최소값이 첫 번째 셀에 있을 때.
            index_no = index_no + 1; orig_no = index_no
            #df_Index[index_no,] = c(linefrom_no, Phase, index_no, i, j-1, Comp.data[1])
            df_temp = data.frame(Line=linefrom_no, Phase=subphase, Index=index_no, RowIndex=i, ColIndex=j-1, Cost=Comp.data[1])
            df_Index = rbind(df_Index, df_temp)
            i = i; j = j-1 
         } else if (which.min(Comp.data) == 2){ # 최소값이 두 번째 셀에 있을 때.
            index_no = index_no + 1; orig_no = index_no
            #df_Index[index_no,] = c(linefrom_no, Phase, index_no, i-1, j-1, Comp.data[2])
            df_temp = data.frame(Line=linefrom_no, Phase=subphase, Index=index_no, RowIndex=i-1, ColIndex=j-1, Cost=Comp.data[2])
            df_Index = rbind(df_Index, df_temp)
            i = i-1; j = j-1
         } else{ # 최소값이 세 번째 셀에 있을 때.
            index_no = index_no + 1; orig_no = index_no
            #df_Index[index_no,] = c(linefrom_no, Phase, index_no, i-1, j, Comp.data[3])
            df_temp = data.frame(Line=linefrom_no, Phase=subphase, Index=index_no, RowIndex=i-1, ColIndex=j, Cost=Comp.data[3])
            df_Index = rbind(df_Index, df_temp)
            i = i-1; j = j
         }
      } else{ # 최소값이 중복되었을 경우.
         repetition = which(Comp.data == Comp.data[which.min(Comp.data)]) #중복된 최소값의 위치를 저장한다.
         orig_no = index_no; u = i; v = j;
         
         for(k in 1:length(repetition)){ #중복된 값의 수 만큼 실행.

            if(k == 1){ #첫 번째 중복이면 실행.
               if(repetition[1] == 1){ #중복된 값의 위치가 1이면 실행. 첫 번째 중복 값은 무조건 df_Index에 저장.
                  index_no = index_no + 1
                  #df_Index[index_no,] = c(linefrom_no, Phase, index_no, i, j-1, Comp.data[1])
                  df_temp = data.frame(Line=linefrom_no, Phase=subphase, Index=index_no, RowIndex=i, ColIndex=j-1, Cost=Comp.data[1])
                  df_Index = rbind(df_Index, df_temp)
                  i = i; j = j-1
               } else{ #중복된 값의 위치가 2이면 실행. 중복이면 3이 될 수 없음.
                  index_no = index_no + 1
                  #df_Index[index_no,] = c(linefrom_no, Phase, index_no, i-1, j-1, Comp.data[2])
                  df_temp = data.frame(Line=linefrom_no, Phase=subphase, Index=index_no, RowIndex=i-1, ColIndex=j-1, Cost=Comp.data[2])
                  df_Index = rbind(df_Index, df_temp)
                  i = i-1; j = j-1
               }
            } else if(repetition[k] == 2){ # 중복된 값 저장
               same_no = same_no + 1; lineto_no = lineto_no + 1
               df_SameIndex[same_no,] = c(linefrom_no, lineto_no, orig_no, same_no, u-1, v-1, Comp.data[2])

            } else { # 중복된 값 저장
               same_no = same_no + 1; lineto_no = lineto_no + 1
               df_SameIndex[same_no,] = c(linefrom_no, lineto_no, orig_no, same_no, u-1, v, Comp.data[3])
            }
         } #for end.
      }#else end
   
   }#while end.{여기까지 완성. 아래 코드는 중복 값에 대한 warping path를 찾는 것.}
   
   lst_Variable = list(df_Index, df_SameIndex)
   names(lst_Variable) = c("df_Index", "df_SameIndex")
   return(lst_Variable)
}


#===list를 인수로 받아 다른 full index를 data.frame 형태로 리턴한다.
ft_OtherWarpingPath <- function(lst_data_df, metrix_data){
   
   #list로 된 데이터를 다시 data.frame으로 변경하여 저장.
   df_Index = lst_data_df$df_Index
   df_SameIndex = lst_data_df$df_SameIndex
   
   cost.matrix = metrix_data
   subphase = "ALL"
   
   
   otherline = 1
   #if(otherline != nrow(df_SameIndex)){#otherline이 존재하는지 확인(중복값이 있는지 확인).
      
      while(otherline <= nrow(df_SameIndex)){
         
         indexCounter = nrow(df_Index)
         # lineto_no = max(df_SameIndex$lineto)
         
         
         start_no = indexCounter + 1 # 첫번째 라인에서 가져온 데이터를 입력할 시작 위치.
         end_no = indexCounter + df_SameIndex[otherline, 3] #may rm 첫 번째 라인에서 가져온 데이터를 입력할 마지막 위치.
         
         copyStart <<- which(df_Index$Line == df_SameIndex[otherline, 1] & df_Index$Index == 1)
         copyEnd = (copyStart + df_SameIndex[otherline, 3]) - 1
         
         
         temp_Copy = df_Index[copyStart:copyEnd, ] #Orig 값까지를 가져온다.
         temp_Copy$Line = df_SameIndex[otherline, 2] #Lineto 값을 입력한다.
         
         df_Index = rbind(df_Index, temp_Copy)
         
         #View(df_Index)
         
         #===================================================================================================================
         i = (df_SameIndex[otherline, 5])
         j = (df_SameIndex[otherline, 6])
         
         
         line_no = (df_SameIndex[otherline, 2]) #LineTo 값을 입력한다.
         index_no = (df_SameIndex[otherline, 3])+1 #oring 값을 입력한다.
         full_Index_no = nrow(df_Index)+1
         full_Same_no = nrow(df_SameIndex)
         linefrom_no = (df_SameIndex[otherline, 2])
         lineto_no = max(df_SameIndex$LineTo)
         same_no = 0
         
         df_temp = data.frame(Line=line_no, Phase=subphase, Index=index_no, RowIndex=i, ColIndex=j, Cost=df_SameIndex[otherline, 7])
         df_Index = rbind(df_Index, df_temp)
         
         # 비교할 데이터를 만드는 sub 함수이다.
         ft_CompData <- function(x, y){
            
            i = x; j = y
            
            if(i > 1 & j > 1){ #X축과 Y축 값이 모두 1 이상인 경우. 즉 왼쪽과 아래쪽에 모두 데이터가 있는 경우.
               CompData = as.vector(c(cost.matrix[i, j-1], cost.matrix[i-1, j-1], cost.matrix[i-1, j])) # 반시계 방향으로 처리함.
            } else if(i == 1 & j > 1){
               CompData = c(cost.matrix[i, j-1], cost.matrix[i, j-1]+1, cost.matrix[i, j-1]+1) #왼쪽값만 입력하고 나머지는 1이 더 큰 값을 입력한다.
            } else if(i > 1 & j == 1){
               CompData = c(cost.matrix[i-1, j]+1, cost.matrix[i-1, j]+1, cost.matrix[i-1, j]) #아래값만 입력하고 나머지는 1이 큰 값을 입력한다.
            } else if(i == 1 & j == 1){ # 없어도 돼지만 귀찮아서 그냥 둠.
               break
            }
            return(CompData)
         }
         
         
         while (i >= 1 | j >= 1) { #맨 위에서부터 처리함.
            # if(i > 1 & j > 1){ #X축과 Y축 값이 모두 1 이상인 경우. 즉 왼쪽과 아래쪽에 모두 데이터가 있는 경우.
            #    Comp.data = c(cost.matrix[i, j-1], cost.matrix[i-1, j-1], cost.matrix[i-1, j]) # 반시계 방향으로 처리함.
            # } else if(i == 1 & j > 1){
            #    Comp.data = c(cost.matrix[i, j-1], cost.matrix[i, j-1]+1, cost.matrix[i, j-1]+1) #왼쪽값만 입력하고 나머지는 1이 더 큰 값을 입력한다.
            # } else if(i > 1 & j == 1){
            #    Comp.data = c(cost.matrix[i-1, j]+1, cost.matrix[i-1, j]+1, cost.matrix[i-1, j]) #아래값만 입력하고 나머지는 1이 큰 값을 입력한다.
            # } else if(i == 1 & j == 1){
            #    break
            # }
            
            if(i == 1 & j == 1){
               break
            }
            
            Comp.data = ft_CompData(i, j)
            
            if(length(which(Comp.data == Comp.data[which.min(Comp.data)])) == 1){ # 최소값이 중복되지 않은 경우.
               if (which.min(Comp.data) == 1){ # 최소값이 첫 번째 셀에 있을 때.
                  index_no = index_no + 1
                  full_Index_no = full_Index_no + 1
                  #df_Index[full_Index_no,] = c(line_no, Phase, index_no, i, j-1, Comp.data[1])
                  df_temp = data.frame(Line=line_no, Phase=subphase, Index=index_no, RowIndex=i, ColIndex=j-1, Cost=Comp.data[1])
                  df_Index = rbind(df_Index, df_temp)
                  i = i; j = j-1 
               } else if (which.min(Comp.data) == 2){ # 최소값이 두 번째 셀에 있을 때.
                  index_no = index_no + 1
                  full_Index_no = full_Index_no + 1
                  #df_Index[full_Index_no,] = c(line_no, Phase, index_no, i-1, j-1, Comp.data[2])
                  df_temp = data.frame(Line=line_no, Phase=subphase, Index=index_no, RowIndex=i-1, ColIndex=j-1, Cost=Comp.data[2])
                  df_Index = rbind(df_Index, df_temp)
                  i = i-1; j = j-1
               } else{ # 최소값이 세 번째 셀에 있을 때.
                  index_no = index_no + 1
                  full_Index_no = full_Index_no + 1
                  #df_Index[full_Index_no,] = c(line_no, Phase, index_no, i-1, j, Comp.data[3])
                  df_temp = data.frame(Line=line_no, Phase=subphase, Index=index_no, RowIndex=i-1, ColIndex=j, Cost=Comp.data[3])
                  df_Index = rbind(df_Index, df_temp)
                  i = i-1; j = j
               }
            } else{ # 최소값이 중복되었을 경우.
               repetition = which(Comp.data == Comp.data[which.min(Comp.data)]) #중복된 최소값의 위치를 저장한다.
               orig_no = index_no; u = i; v = j
               
               for(k in 1:length(repetition)){ #중복된 값의 수 만큼 실행.
                  
                  if(k == 1){ #첫 번째 중복이면 실행.
                     if(repetition[1] == 1){ #중복된 값의 위치가 1이면 실행. 첫 번째 중복 값은 무조건 df_Index에 저장.
                        index_no = index_no + 1
                        full_Index_no = full_Index_no + 1
                        #df_Index[full_Index_no,] = c(line_no, Phase, index_no, i, j-1, Comp.data[1])
                        df_temp = data.frame(Line=line_no, Phase=subphase, Index=index_no, RowIndex=i, ColIndex=j-1, Cost=Comp.data[1])
                        df_Index = rbind(df_Index, df_temp)
                        i = i; j = j-1
                     } else{ #중복된 값의 위치가 2이면 실행. 중복이면 3이 될 수 없음.
                        index_no = index_no + 1
                        full_Index_no = full_Index_no + 1
                        #df_Index[full_Index_no,] = c(line_no, Phase, index_no, i-1, j-1, Comp.data[2])
                        df_temp = data.frame(Line=line_no, Phase=subphase, Index=index_no, RowIndex=i-1, ColIndex=j-1, Cost=Comp.data[2])
                        df_Index = rbind(df_Index, df_temp)
                        i = i-1; j = j-1
                     }
                  } else if(repetition[k] == 2){ # 중복된 값 저장
                     full_Same_no = full_Same_no + 1;  lineto_no = lineto_no + 1; same_no = same_no +1
                     df_SameIndex[full_Same_no,] = c(linefrom_no, lineto_no, orig_no, same_no, u-1, v-1, Comp.data[2])
                     
                  } else { # 중복된 값 저장
                     full_Same_no = full_Same_no + 1;  lineto_no = lineto_no + 1; same_no = same_no +1
                     df_SameIndex[full_Same_no,] = c(linefrom_no, lineto_no, orig_no, same_no, u-1, v, Comp.data[3])
                  }
               } #for end.
            }#else end
         }#while end level 2.
         
         otherline = otherline + 1
         
      }#while end level 1.
   #}#if end level 1.
   
   return(df_Index)
   #lis.WarpingPath <<- df_Index #글로벌 변수에 입력
} #function end.


#===a, b 두 개의 벡터를 입력으로 받아, 두 벡터 사이의 각을 계산하여 리턴함.
ft_DotProduct <- function(a, b){
   # dot_value = a%*%b 
   # norm_a = norm(a, type="2")
   # norm_b = norm(b, type="2")
   
   up = a%*%b
   down = len(a) * len(b)
   
   if (up > down){
      up = down
   }
   
   theta_radian = acos(up / down)
   theta_angle = (theta_radian*180)/pi
   as.numeric(theta_angle)
   
   # theta_radian = acos(dot_value / (norm_a * norm_b))
   # theta_angle = (theta_radian*180)/pi
   # as.numeric(theta_angle)
   return(theta_angle)
}


#===dataframe을 입력으로 받아 rmse를 계산하여 dataframe으로 리턴한다.
#===df_full_path를 입력받아 df_rmse를 리턴한다.
ft_RMSE <- function(df_data_df){
   
   df_rmse = data.frame(rmse=numeric())
   counter = max(df_data_df$Line) #warping path의 개수를 알아낸다.
   
   for(i in 1:counter){ # Line 수 만큼 처리.
      df_analydata = subset(df_data_df, Line == i)
      index_counter = max(df_analydata$Index)
      temp_rmse = data.frame(rmse=numeric())
      
      #X, Y 축을 설정하기 위해서 Row와 Column 순서를 맞춤.
      diagonal = c(max(df_analydata$ColIndex), max(df_analydata$RowIndex))
      
      #임시용#여기에 방정식을 입력
      #incline = diagonal[2]/diagonal[1]
      #max_deviation = diagonal[1]*diagonal[2]
      
      for(ii in 1:index_counter){ # index 수 만큼 실행
         spot = c(df_analydata[ii, 5], df_analydata[ii, 4])
         
         spot_length = sqrt((df_analydata[ii, 5]^2) + (df_analydata[ii, 4])^2)
         
         #제작한 ft_DotProduct() 함수 대신에 matlib 팩키지에 있는 angle 함수를 사용
         angle = ft_DotProduct(diagonal, spot)
         #angle = matlib::angle(diagonal, spot)
         
         # if(ii == 1){
         #    print(diagonal); print(spot) ; print(angle)
         # }
         
         radian = (angle*pi)/180
         temp_rmse[ii,1] = spot_length * sin(radian)
      }#2nd for
      
      df_rmse = rbind(df_rmse, temp_rmse)
      
   }#1th for
   
   return(df_rmse)
}


#===최종 분석 대상인 df_full_path에 대한 분석을 수행
ft_analysis <- function(df_data_df, n1, n2){
   
   df_full_path = df_data_df
   
   #결과를 저장할 data.frame 생성
   df_analysis = data.frame(Path_No=numeric(), Index1_length=numeric(), Index2_length=numeric(), Path_Index=numeric(), Sum_of_Cost=numeric(), Cost_div_N=numeric(),
                            Cost_div_diagonal=numeric(), sum_of_RMSE=numeric(), RMSE_div_N=numeric(), RMSE_div_diagonal=numeric(), Mean_of_RMSE=numeric(),
                            RMSE_SD=numeric())
   
   #warping path 개수를 저장
   path_count = max(df_full_path$Line)
   
   for(i in 1:path_count){
      df_temp_line=subset(df_full_path, Line == i)
      index_count = max(df_temp_line$Index)
      cost_sum = sum(df_temp_line$Cost)
      cost_div_n = sum(df_temp_line$Cost) / (n1 + n2)
      cost_div_diag = sum(df_temp_line$Cost) / sqrt(n1^2 + n2^2)
      RMSE_sum = sum(df_temp_line$rmse)
      RMSE_div_n = sum(df_temp_line$rmse) / (n1 + n2)
      RMSE_div_diag = sum(df_temp_line$rmse) / sqrt(n1^2 + n2^2)
      rmse_mean_value = mean(df_temp_line$rmse)
      rmse_sd = sd(df_temp_line$rmse)
      
      df_save = data.frame(Path_No=i, Index1_length=n1, Index2_length=n2, Path_Index=index_count, Sum_of_Cost=cost_sum, Cost_div_N = cost_div_n,
                           Cost_div_diagonal=cost_div_diag, sum_of_RMSE=RMSE_sum, RMSE_div_N=RMSE_div_n, RMSE_div_diagonal=RMSE_div_diag, Mean_of_RMSE=rmse_mean_value,
                           RMSE_SD = rmse_sd)
      df_analysis = rbind(df_analysis, df_save)
   }
   
   return(df_analysis)
}


#===분석 결과를 그래프로 그리는 함수: df_full_path, analysis_summary, df_refquery, df_warpingpath를 입력으로 받는다.
ft_graph <- function(df_data1_lst, data2, df_data3, df_data4, df_data5){
   
   df_full_path = df_data1_lst; analysis_summary = data2;  df_refquery = df_data3;  df_warpingpath = df_data4
   
   #분석에 사용된 raw data 출력
   # NA 값을 제거하기 위해서 임시 저장한다.
   temp.index1 = df_refquery$index1
   temp.index2 = df_refquery$index2
   
   # NA 값을 제거한다.
   rawdata_index1 = temp.index1[!is.na(temp.index1)]
   rawdata_index2 = temp.index2[!is.na(temp.index2)]
   
   ref_length = 1:length(rawdata_index1)
   query_length = 1:length(rawdata_index2)
   length_max = max(length(rawdata_index1), length(rawdata_index2))
   
   #Raw data 출력
   palette_rawdata = c("index1"="#000000", "index2"="#0000ff")
   graph_rawdata <- ggplot() +
      geom_point(aes(x=ref_length, y=rawdata_index1, colour = "index1"), size = 0.2) +
      geom_point(aes(x=query_length, y=rawdata_index2, colour ="index2"), size = 0.2) +
      geom_hline(yintercept = 0, color = "black") +
      scale_color_manual(values = palette_rawdata) +
      xlab("Frames") +
      ylab("Values") +
      #coord_fixed(ratio = 1) +
      scale_x_continuous(breaks = seq(0, length_max, 10)) +
      scale_y_continuous(breaks = round(seq(min(rawdata_index1), max(rawdata_index1), length.out = 11), digits = 2)) +
      ggtitle("Graph of Raw data") +
      theme(plot.title = element_text(hjust = 0.5)) +
      theme(legend.title = element_blank()) +
      theme_bw()
   
     
   #모든 path 출력
   #palette_rawdata = c("index1"="#000000", "index2"="#0000ff", "axis"="#797980", "red"="#ff0000")
   graph_all_path <- ggplot(data=df_full_path, aes(x=ColIndex, y=RowIndex, group = Line, color = Line)) +
      geom_line(size = 0.2) +
      geom_point(size = 0.2)+
      geom_segment(x=1, xend=max(df_full_path$ColIndex), y=1, yend=max(df_full_path$RowIndex), color = "red", size = 0.2) +
      theme_bw()
   
   #Optimal warping path 출력
   graph_warpingpath <- ggplot(data = df_warpingpath, aes(x=ColIndex, y=RowIndex)) +
      geom_line(size=0.2) + geom_point(size=0.2) + geom_segment(x=1, xend=max(df_warpingpath$ColIndex), y=1, yend=max(df_warpingpath$RowIndex), color = "red", size = 0.2) +
      theme_bw()
   
   
   #Raw data에서의 matching 그래프 출력
   palette_rawdata = c("index1"="#000000", "index2"="#0000ff", "axis"="#797980", "red"="#ff0000")
   graph_matching <- ggplot() +
      geom_point(aes(x=ref_length, y=rawdata_index1, colour = "index1"), size = 0.2) +
      geom_point(aes(x=query_length, y=rawdata_index2, colour ="index2"), size = 0.2) +
      geom_hline(yintercept = 0, color = "black") +
      scale_color_manual(values = palette_rawdata) +
      xlab("Frames") +
      ylab("Values") +
      #coord_fixed(ratio = 1) +
      scale_x_continuous(breaks = seq(0, length_max, 10)) +
      scale_y_continuous(breaks = round(seq(min(rawdata_index1), max(rawdata_index1), length.out = 11), digits = 2)) +
      ggtitle("Matching Graph for Raw data") +
      theme_bw()
   
   for (i in 1:length(df_warpingpath$Index)){
      graph_matching <- graph_matching +
         annotate("segment", x = df_warpingpath$ColIndex[i], xend = df_warpingpath$RowIndex[i],
                  y = rawdata_index2[df_warpingpath$ColIndex[i]], yend = rawdata_index1[df_warpingpath$RowIndex[i]], color = "red", size = 0.2)
   }
   
   graph = list()
   graph$rawdata = graph_rawdata
   graph$allpath = graph_all_path
   graph$warpingpath = graph_warpingpath
   graph$matching = graph_matching
   return(graph)
}

#===Total.Results를 엑셀 파일로 저장한다.(그래프는 미포함.)
ft_Save_TotalResults <- function(lst_data_non, target = 0){
   # target은 저장 대상을 의미함. 0인 경우 디폴트, 1인 경우 phase별 분석 결과임.
   
   Total.Results = lst_data_non
   # print("target")
   # print(target)
   
   if(target == 0){
      #Workbook 생성
      TR = createWorkbook("Total_Results")
   
      #생성된 Workbook에 sheet를 생성.
      addWorksheet(TR, "Raw_Data")
      addWorksheet(TR, "Cost_Matrix")
      addWorksheet(TR, "Full_Path")
      addWorksheet(TR, "Analysis_Summary")
      addWorksheet(TR, "Warping_Path")
   
      writeData(TR, "Raw_Data", Total.Results$index1, colNames = TRUE, rowNames = FALSE, startCol = "A", startRow = 1)
      writeData(TR, "Raw_Data", Total.Results$index2, colNames = TRUE, rowNames = FALSE, startCol = "B", startRow = 1)
      writeData(TR, "Cost_Matrix", Total.Results$costmatrix, colNames = TRUE, rowNames = TRUE, startCol = "A", startRow = 1)
      writeData(TR, "Full_Path", Total.Results$fullpath, colNames = TRUE, rowNames = FALSE, startCol = "A", startRow = 1)
      writeData(TR, "Analysis_Summary", Total.Results$analysis, colNames = TRUE, rowNames = TRUE, startCol = "A", startRow = 1)
      writeData(TR, "Warping_Path", Total.Results$warpingpath, colNames = TRUE, rowNames = FALSE, startCol = "A", startRow = 1)
   } else if(target ==1){
      #Workbook 생성
      TR = createWorkbook("Total_Results_phase")
      
      #생성된 Workbook에 sheet를 생성.
      addWorksheet(TR, "Pre_Raw_Data")
      addWorksheet(TR, "Pre_Cost_Matrix")
      addWorksheet(TR, "Pre_Full_Path")
      addWorksheet(TR, "Pre_Analysis")
      addWorksheet(TR, "Pre_Warping_Path")
      addWorksheet(TR, "Post_Raw_Data")
      addWorksheet(TR, "Post_Cost_Matrix")
      addWorksheet(TR, "Post_Full_Path")
      addWorksheet(TR, "Post_Analysis")
      addWorksheet(TR, "Post_Warping_Path")
      addWorksheet(TR, "Merged_Warping_Path")
      addWorksheet(TR, "Merged_Analysis")
      
      writeData(TR, "Pre_Raw_Data", Total.Results$pre_index1, colNames = TRUE, rowNames = FALSE, startCol = "A", startRow = 1)
      writeData(TR, "Pre_Raw_Data", Total.Results$pre_index2, colNames = TRUE, rowNames = FALSE, startCol = "B", startRow = 1)
      writeData(TR, "Pre_Cost_Matrix", Total.Results$pre_costmatrix, colNames = TRUE, rowNames = TRUE, startCol = "A", startRow = 1)
      writeData(TR, "Pre_Full_Path", Total.Results$pre_fullpath, colNames = TRUE, rowNames = FALSE, startCol = "A", startRow = 1)
      writeData(TR, "Pre_Analysis", Total.Results$pre_analysis, colNames = TRUE, rowNames = TRUE, startCol = "A", startRow = 1)
      writeData(TR, "Pre_Warping_Path", Total.Results$pre_warpingpath, colNames = TRUE, rowNames = FALSE, startCol = "A", startRow = 1)
      writeData(TR, "Post_Raw_Data", Total.Results$post_index1, colNames = TRUE, rowNames = FALSE, startCol = "A", startRow = 1)
      writeData(TR, "Post_Raw_Data", Total.Results$post_index2, colNames = TRUE, rowNames = FALSE, startCol = "B", startRow = 1)
      writeData(TR, "Post_Cost_Matrix", Total.Results$post_costmatrix, colNames = TRUE, rowNames = TRUE, startCol = "A", startRow = 1)
      writeData(TR, "Post_Full_Path", Total.Results$post_fullpath, colNames = TRUE, rowNames = FALSE, startCol = "A", startRow = 1)
      writeData(TR, "Post_Analysis", Total.Results$post_analysis, colNames = TRUE, rowNames = TRUE, startCol = "A", startRow = 1)
      writeData(TR, "Post_Warping_Path", Total.Results$post_warpingpath, colNames = TRUE, rowNames = FALSE, startCol = "A", startRow = 1)
      writeData(TR, "Merged_Warping_Path", Total.Results$MergedWarpingpath, colNames = TRUE, rowNames = FALSE, startCol = "A", startRow = 1)
      writeData(TR, "Merged_Analysis", Total.Results$MergedAnalysis, colNames = TRUE, rowNames = FALSE, startCol = "A", startRow = 1)
   }
   
   #폴더를 생성한다.
   folder = file.path(getwd(), "Total_Results")
   if(dir.exists(folder) == FALSE){
      dir.create(folder)
   }
   
   if(target == 0){
      #파일을 생성한다.
      filecounter = 1
      while(filecounter > 0){
         fname = paste0(folder, "/Total_Results_", sprintf("%03d", filecounter), ".XLSX")
         if(file.exists(fname) == FALSE){
            saveWorkbook(TR, file = fname)
            #file.create(fname)
            break
         }
         filecounter = filecounter + 1
      }
   } else if(target == 1){
      #파일을 생성한다.
      filecounter = 1
      while(filecounter > 0){
         fname = paste0(folder, "/Total_Results_Phase_", sprintf("%03d", filecounter), ".XLSX")
         if(file.exists(fname) == FALSE){
            saveWorkbook(TR, file = fname)
            #file.create(fname)
            break
         }
         filecounter = filecounter + 1
      }
   }
   
   
   #엑셀 파일을 저장한다.
   #saveWorkbook(TR, file = fname)
   
   #파일 경로와 파일이름을 분리한 후, 확장자를 제외한 파일명만 fname_graph에 저장한다.
   dirpath = dirname(fname)
   filepart = basename(fname)
   fname_graph = file_path_sans_ext(filepart)
   
   if(target == 0){
      print("시간이 걸림: <<완료됨.>>이 나타날 때까지 기다리세요.: 1/4 진행 중")
      fname_graphrawdata = paste0(fname_graph, "_graph_Rawdata.png")
      ggsave(plot = Total.Results$graph_rawdata, filename = fname_graphrawdata, path = dirpath, dpi = 300 )
   
      print("시간이 걸림: <<완료됨.>>이 나타날 때까지 기다리세요.: 2/4 진행 중")
      fname_graphallpath = paste0(fname_graph, "_graph_Allpath.png")
      ggsave(plot = Total.Results$graph_allpath, filename = fname_graphallpath, path = dirpath, dpi = 300 )
   
      print("시간이 걸림: <<완료됨.>>이 나타날 때까지 기다리세요.: 3/4 진행 중")
      fname_graphwarping = paste0(fname_graph, "_graph_Warpingpath.png")
      ggsave(plot = Total.Results$graph_warpingpath, filename = fname_graphwarping, path = dirpath, dpi = 300 )
   
      print("시간이 걸림: <<완료됨.>>이 나타날 때까지 기다리세요.: 4/4 진행 중")
      fname_graphmatchig = paste0(fname_graph, "_graph_Matching.png")
      ggsave(plot = Total.Results$graph_matching, filename = fname_graphmatchig, path = dirpath, dpi = 300 )
      print("<<완료됨!!.>>")
   } else if(target == 1){
      print("시간이 걸림: <<완료됨.>>이 나타날 때까지 기다리세요.: 1/9 진행 중")
      fname_graphrawdata = paste0(fname_graph, "_Contact_graph_Rawdata.png")
      ggsave(plot = Total.Results$pre_graph_rawdata, filename = fname_graphrawdata, path = dirpath, dpi = 300 )
      
      print("시간이 걸림: <<완료됨.>>이 나타날 때까지 기다리세요.: 2/9 진행 중")
      fname_graphallpath = paste0(fname_graph, "_Contact_graph_Allpath.png")
      ggsave(plot = Total.Results$pre_graph_allpath, filename = fname_graphallpath, path = dirpath, dpi = 300 )
      
      print("시간이 걸림: <<완료됨.>>이 나타날 때까지 기다리세요.: 3/9 진행 중")
      fname_graphwarping = paste0(fname_graph, "_Contact_graph_Warpingpath.png")
      ggsave(plot = Total.Results$pre_graph_warpingpath, filename = fname_graphwarping, path = dirpath, dpi = 300 )
      
      print("시간이 걸림: <<완료됨.>>이 나타날 때까지 기다리세요.: 4/9 진행 중")
      fname_graphmatchig = paste0(fname_graph, "_Contact_graph_Matching.png")
      ggsave(plot = Total.Results$pre_graph_matching, filename = fname_graphmatchig, path = dirpath, dpi = 300 )
      
      print("시간이 걸림: <<완료됨.>>이 나타날 때까지 기다리세요.: 5/9 진행 중")
      fname_graphrawdata = paste0(fname_graph, "_Swing_graph_Rawdata.png")
      ggsave(plot = Total.Results$post_graph_rawdata, filename = fname_graphrawdata, path = dirpath, dpi = 300 )
      
      print("시간이 걸림: <<완료됨.>>이 나타날 때까지 기다리세요.: 6/9 진행 중")
      fname_graphallpath = paste0(fname_graph, "_Swing_graph_Allpath.png")
      ggsave(plot = Total.Results$post_graph_allpath, filename = fname_graphallpath, path = dirpath, dpi = 300 )
      
      print("시간이 걸림: <<완료됨.>>이 나타날 때까지 기다리세요.: 7/9 진행 중")
      fname_graphwarping = paste0(fname_graph, "_Swing_graph_Warpingpath.png")
      ggsave(plot = Total.Results$post_graph_warpingpath, filename = fname_graphwarping, path = dirpath, dpi = 300 )
      
      print("시간이 걸림: <<완료됨.>>이 나타날 때까지 기다리세요.: 8/9 진행 중")
      fname_graphmatchig = paste0(fname_graph, "_Swing_graph_Matching.png")
      ggsave(plot = Total.Results$post_graph_matching, filename = fname_graphmatchig, path = dirpath, dpi = 300 )
      
      print("시간이 걸림: <<완료됨.>>이 나타날 때까지 기다리세요.: 9/9 진행 중")
      fname_graphmergedmatchig = paste0(fname_graph, "_Merged_graph_Matching.png")
      ggsave(plot = Total.Results$Merged_graph_Matching, filename = fname_graphmergedmatchig, path = dirpath, dpi = 300 )
      print("<<완료됨!!.>>")
      
   }
 
}

#===Phase 별로 분석하였을 때 그 결과를 저장하는 함수이다. ft_Save_TotalRreults() 함수의 일부를 따로 때어서 만들었다. 
ft_Save_phase <- function(lst_data_non){
   
   
   Total.Results = lst_data_non
   
   #Workbook 생성
   TR = createWorkbook("Total_Results_phase")
      
   #생성된 Workbook에 sheet를 생성.
   addWorksheet(TR, "Pre_Raw_Data")
   addWorksheet(TR, "Pre_Cost_Matrix")
   addWorksheet(TR, "Pre_Full_Path")
   addWorksheet(TR, "Pre_Analysis")
   addWorksheet(TR, "Pre_Warping_Path")
   addWorksheet(TR, "Post_Raw_Data")
   addWorksheet(TR, "Post_Cost_Matrix")
   addWorksheet(TR, "Post_Full_Path")
   addWorksheet(TR, "Post_Analysis")
   addWorksheet(TR, "Post_Warping_Path")
   addWorksheet(TR, "Merged_Warping_Path")
   addWorksheet(TR, "Merged_Analysis")
      
   writeData(TR, "Pre_Raw_Data", Total.Results$pre_index1, colNames = TRUE, rowNames = FALSE, startCol = "A", startRow = 1)
   writeData(TR, "Pre_Raw_Data", Total.Results$pre_index2, colNames = TRUE, rowNames = FALSE, startCol = "B", startRow = 1)
   writeData(TR, "Pre_Cost_Matrix", Total.Results$pre_costmatrix, colNames = TRUE, rowNames = TRUE, startCol = "A", startRow = 1)
   writeData(TR, "Pre_Full_Path", Total.Results$pre_fullpath, colNames = TRUE, rowNames = FALSE, startCol = "A", startRow = 1)
   writeData(TR, "Pre_Analysis", Total.Results$pre_analysis, colNames = TRUE, rowNames = TRUE, startCol = "A", startRow = 1)
   writeData(TR, "Pre_Warping_Path", Total.Results$pre_warpingpath, colNames = TRUE, rowNames = FALSE, startCol = "A", startRow = 1)
   writeData(TR, "Post_Raw_Data", Total.Results$post_index1, colNames = TRUE, rowNames = FALSE, startCol = "A", startRow = 1)
   writeData(TR, "Post_Raw_Data", Total.Results$post_index2, colNames = TRUE, rowNames = FALSE, startCol = "B", startRow = 1)
   writeData(TR, "Post_Cost_Matrix", Total.Results$post_costmatrix, colNames = TRUE, rowNames = TRUE, startCol = "A", startRow = 1)
   writeData(TR, "Post_Full_Path", Total.Results$post_fullpath, colNames = TRUE, rowNames = FALSE, startCol = "A", startRow = 1)
   writeData(TR, "Post_Analysis", Total.Results$post_analysis, colNames = TRUE, rowNames = TRUE, startCol = "A", startRow = 1)
   writeData(TR, "Post_Warping_Path", Total.Results$post_warpingpath, colNames = TRUE, rowNames = FALSE, startCol = "A", startRow = 1)
   writeData(TR, "Merged_Warping_Path", Total.Results$MergedWarpingpath, colNames = TRUE, rowNames = FALSE, startCol = "A", startRow = 1)
   writeData(TR, "Merged_Analysis", Total.Results$MergedAnalysis, colNames = TRUE, rowNames = FALSE, startCol = "A", startRow = 1)
   
   
   #폴더를 생성한다.
   folder = file.path(getwd(), "Total_Results")
   if(dir.exists(folder) == FALSE){
      dir.create(folder)
   }
   
   #파일을 생성한다.
   filecounter = 1
   while(filecounter > 0){
      fname = paste0(folder, "/Total_Results_Phase_", sprintf("%03d", filecounter), ".XLSX")
      if(file.exists(fname) == FALSE){
         saveWorkbook(TR, file = fname)
         #file.create(fname)
         break
      }
      filecounter = filecounter + 1
   }
   
   
   #엑셀 파일을 저장한다.
   #saveWorkbook(TR, file = fname)
   
   #파일 경로와 파일이름을 분리한 후, 확장자를 제외한 파일명만 fname_graph에 저장한다.
   dirpath = dirname(fname)
   filepart = basename(fname)
   fname_graph = file_path_sans_ext(filepart)
   
   print("시간이 걸림: <<완료됨.>>이 나타날 때까지 기다리세요.: 1/10 진행 중")
   fname_graphrawdata = paste0(fname_graph, "_Contact_graph_Rawdata.png")
   ggsave(plot = Total.Results$pre_graph_rawdata, filename = fname_graphrawdata, path = dirpath, dpi = 300 )
      
   print("시간이 걸림: <<완료됨.>>이 나타날 때까지 기다리세요.: 2/10 진행 중")
   fname_graphallpath = paste0(fname_graph, "_Contact_graph_Allpath.png")
   ggsave(plot = Total.Results$pre_graph_allpath, filename = fname_graphallpath, path = dirpath, dpi = 300 )
      
   print("시간이 걸림: <<완료됨.>>이 나타날 때까지 기다리세요.: 3/10 진행 중")
   fname_graphwarping = paste0(fname_graph, "_Contact_graph_Warpingpath.png")
   ggsave(plot = Total.Results$pre_graph_warpingpath, filename = fname_graphwarping, path = dirpath, dpi = 300 )
      
   print("시간이 걸림: <<완료됨.>>이 나타날 때까지 기다리세요.: 4/10 진행 중")
   fname_graphmatchig = paste0(fname_graph, "_Contact_graph_Matching.png")
   ggsave(plot = Total.Results$pre_graph_matching, filename = fname_graphmatchig, path = dirpath, dpi = 300 )
      
   print("시간이 걸림: <<완료됨.>>이 나타날 때까지 기다리세요.: 5/10 진행 중")
   fname_graphrawdata = paste0(fname_graph, "_Swing_graph_Rawdata.png")
   ggsave(plot = Total.Results$post_graph_rawdata, filename = fname_graphrawdata, path = dirpath, dpi = 300 )
      
   print("시간이 걸림: <<완료됨.>>이 나타날 때까지 기다리세요.: 6/10 진행 중")
   fname_graphallpath = paste0(fname_graph, "_Swing_graph_Allpath.png")
   ggsave(plot = Total.Results$post_graph_allpath, filename = fname_graphallpath, path = dirpath, dpi = 300 )
      
   print("시간이 걸림: <<완료됨.>>이 나타날 때까지 기다리세요.: 7/10 진행 중")
   fname_graphwarping = paste0(fname_graph, "_Swing_graph_Warpingpath.png")
   ggsave(plot = Total.Results$post_graph_warpingpath, filename = fname_graphwarping, path = dirpath, dpi = 300 )
      
   print("시간이 걸림: <<완료됨.>>이 나타날 때까지 기다리세요.: 8/10 진행 중")
   fname_graphmatchig = paste0(fname_graph, "_Swing_graph_Matching.png")
   ggsave(plot = Total.Results$post_graph_matching, filename = fname_graphmatchig, path = dirpath, dpi = 300 )
      
   print("시간이 걸림: <<완료됨.>>이 나타날 때까지 기다리세요.: 9/10 진행 중")
   fname_graphmergedWarping = paste0(fname_graph, "_Merged_graph_Warpingpath.png")
   ggsave(plot = Total.Results$Merged_graph_Warpingpath, filename = fname_graphmergedWarping, path = dirpath, dpi = 300 )
   
   print("시간이 걸림: <<완료됨.>>이 나타날 때까지 기다리세요.: 10/10 진행 중")
   fname_graphmergedmatchig = paste0(fname_graph, "_Merged_graph_Matching.png")
   ggsave(plot = Total.Results$Merged_graph_Matching, filename = fname_graphmergedmatchig, path = dirpath, dpi = 300 )
   print("<<완료됨!!.>>")
      
}

ft_Save_Default_Multi <- function(lst_data_non, target = 0){
   # target은 저장 대상을 의미함. 0인 경우 디폴트, 1인 경우 phase별 분석 결과임.
   
   Total.Results = lst_data_non
   
   #Total.Results의 개체 갯수를 확인.
   counter = length(Total.Results)
   
   #Workbook 생성
   TR = createWorkbook("Total_Results")
      
   #생성된 Workbook에 sheet를 생성.
   addWorksheet(TR, "Cost")
   addWorksheet(TR, "RMSE")
   addWorksheet(TR, "Analysis")
   
   k = 1; m = 1
   for(i in 1:(counter/2)){
      writeData(TR, "Cost", Total.Results[[k]][1], colNames = TRUE, rowNames = FALSE, startCol = i, startRow = 1)
      writeData(TR, "RMSE", Total.Results[[k]][2], colNames = TRUE, rowNames = FALSE, startCol = i, startRow = 1)
      
      writeData(TR, "Analysis", Total.Results[[k+1]], colNames = TRUE, rowNames = TRUE, startCol = "A", startRow = m)
      
      k=k+2; m=m+10
   }
   
   # writeData(TR, "Raw_Data", Total.Results$index1, colNames = TRUE, rowNames = FALSE, startCol = "A", startRow = 1)
   # writeData(TR, "Raw_Data", Total.Results$index2, colNames = TRUE, rowNames = FALSE, startCol = "B", startRow = 1)
   # writeData(TR, "Cost_Matrix", Total.Results$costmatrix, colNames = TRUE, rowNames = TRUE, startCol = "A", startRow = 1)
   # writeData(TR, "Full_Path", Total.Results$fullpath, colNames = TRUE, rowNames = FALSE, startCol = "A", startRow = 1)
   # writeData(TR, "Analysis_Summary", Total.Results$analysis, colNames = TRUE, rowNames = TRUE, startCol = "A", startRow = 1)
   # writeData(TR, "Warping_Path", Total.Results$warpingpath, colNames = TRUE, rowNames = FALSE, startCol = "A", startRow = 1)
   
   #폴더를 생성한다.
   folder = file.path(getwd(), "Total_Results")
   if(dir.exists(folder) == FALSE){
      dir.create(folder)
   }
   
   #파일을 생성한다.
   filecounter = 1
   while(filecounter > 0){
      fname = paste0(folder, "/Total_Results_Multi", sprintf("%03d", filecounter), ".XLSX")
      if(file.exists(fname) == FALSE){
         saveWorkbook(TR, file = fname)
         #file.create(fname)
         break
      }
      filecounter = filecounter + 1
   }
    
   print("파일 생성 완료됨.")
}

#===
ft_Save_Phase_Multi <- function(lst_data_non, target = 0){
   # target은 저장 대상을 의미함. 0인 경우 디폴트, 1인 경우 phase별 분석 결과임.
   
   Total.Results = lst_data_non
   
   #Total.Results의 개체 갯수를 확인.
   counter = length(Total.Results)
   
   #Workbook 생성
   TR = createWorkbook("Total_Results")
   
   #생성된 Workbook에 sheet를 생성.
   addWorksheet(TR, "Cost")
   addWorksheet(TR, "RMSE")
   addWorksheet(TR, "Analysis")
   
   k = 1; m = 1
   for(i in 1:(counter/2)){
      writeData(TR, "Cost", Total.Results[[k]][1], colNames = TRUE, rowNames = FALSE, startCol = i, startRow = 1)
      writeData(TR, "RMSE", Total.Results[[k]][2], colNames = TRUE, rowNames = FALSE, startCol = i, startRow = 1)
      
      writeData(TR, "Analysis", Total.Results[[k+1]], colNames = TRUE, rowNames = TRUE, startCol = "A", startRow = m)
      
      k=k+2; m=m+5
   }
   
   # writeData(TR, "Raw_Data", Total.Results$index1, colNames = TRUE, rowNames = FALSE, startCol = "A", startRow = 1)
   # writeData(TR, "Raw_Data", Total.Results$index2, colNames = TRUE, rowNames = FALSE, startCol = "B", startRow = 1)
   # writeData(TR, "Cost_Matrix", Total.Results$costmatrix, colNames = TRUE, rowNames = TRUE, startCol = "A", startRow = 1)
   # writeData(TR, "Full_Path", Total.Results$fullpath, colNames = TRUE, rowNames = FALSE, startCol = "A", startRow = 1)
   # writeData(TR, "Analysis_Summary", Total.Results$analysis, colNames = TRUE, rowNames = TRUE, startCol = "A", startRow = 1)
   # writeData(TR, "Warping_Path", Total.Results$warpingpath, colNames = TRUE, rowNames = FALSE, startCol = "A", startRow = 1)
   
   #폴더를 생성한다.
   folder = file.path(getwd(), "Total_Results")
   if(dir.exists(folder) == FALSE){
      dir.create(folder)
   }
   
   #파일을 생성한다.
   filecounter = 1
   while(filecounter > 0){
      fname = paste0(folder, "/Total_Results_Phase_Multi", sprintf("%03d", filecounter), ".XLSX")
      if(file.exists(fname) == FALSE){
         saveWorkbook(TR, file = fname)
         #file.create(fname)
         break
      }
      filecounter = filecounter + 1
   }
   
   print("파일 생성 완료됨.")
}


#===DTW가 완료된 main 함수이다.
main_dtw = function(df_data, method=0, process = 0, reference=1, query=2){
   #method=0 : 데이터의 첫번째 열을 reference로 두번째 열을 query로 인식하여 처리한다. Default임. 
   #method=1 : 데이터를 stance phase와 swing phase로 나누어 처리한 후 합치는 방식임.
   #process=0 : 첫 번째 열과 두 번째 열만 분석 대상으로 한다. 
   #process=1 : 데이터 내에 비교 가능한 모든 열을 비교한다. 
   
   #임시로 부여한 것임.
   #reference = 1; query = 2; method = 0
   
   #엑셀 데이터를 읽어들여 data.frame을 리턴한다. 두 문장 중 하나 제거.
   # df_rawdata = ft_ReadExcel()
   df_rawdata = df_data
   
   #print(method); print(process); print(reference); print(query)
   
   #분석에 필요한 열만을 df_refquery에 저장한다. reference와 query에 지정한 열만을 추출하는 과정임.
   #NA 값은 제거되지 않은 상태임.
   if(process == 0){ # 이 부분은 중복된 것으로 파악됨. 
      vt_list = c(paste0("index",reference), paste0("index", query))
      df_refquery = subset(df_rawdata, select=c(vt_list))
   }
   
   #Start----------------------------------------------------
   #결과 모음 리스트인 lst_Total.Results에 rawdata를 입력한다.
   Total.Results = list()
   
   #분석에 사용될 데이터에서 NA 값을 제거하기 위한 작업임.
   temp.ref = df_refquery[, 1] ; temp.query = df_refquery[, 2]
   ref.data = temp.ref[!is.na(temp.ref)]
   query.data = temp.query[!is.na(temp.query)]
   
   #NA가 제거된 자료를 벡터 형태로 Total.Results에 저장한다.
   Total.Results$index1 = data.frame("index1" = ref.data)
   Total.Results$index2 = data.frame("index2" = query.data)
   #End======================================================
   
   #cost matrix를 만들어 mt_cost에 저장한다.
   mt_cost = ft_CostMatrix(df_refquery)
   
   Total.Results$costmatrix = mt_cost
   
   #mt_cost를 보내 첫 번째 warping path를 찾는다.
   lst_warping_path = ft_FirstWarpingPath(mt_cost)
   
   #첫번째 이외의 다른 warping path를 찾아 df_full_path로 저장한다.
   df_full_path = ft_OtherWarpingPath(lst_warping_path, mt_cost)
   
   #RMSE를 계산하여 df_rmse에 저장한다.
   df_rmse = ft_RMSE(df_full_path)
   
   #위에서 구해진 df_rmse를 df_full_path와 합친다.
   df_full_path = cbind(df_full_path, df_rmse)
   
   #Start--------------------------------------------------------------------------------------------
   #Total.Results에 보관하기 위해 df_full_path의 구조를 수정한다.
   #index 값이 큰 값부터 저장된 것을 작은 값부터 저장하는 형태로 변경한다.
   df_full_path = arrange(.data = df_full_path, Line, RowIndex, ColIndex)
   
   #df_full_path$Index 값을 정렬하는 것임.
   for(i in 1:max(df_full_path$Line)){ 
      index_starter = which(df_full_path$Line == i & df_full_path$RowIndex == 1 & df_full_path$ColIndex == 1)
      index_counter = nrow(df_full_path %>% filter(Line == i))
      df_full_path$Index[index_starter:(index_starter + index_counter - 1)] = c(1:index_counter)
   }
   
   Total.Results$fullpath = df_full_path
   #End==============================================================================================
   
   #df_full_path를 분석하기 위해 수행: 분석에 사용된 자료의 길이를 계산.
   index1_n = length(ref.data)
   index2_n = length(query.data)
   
   analysis_summary = ft_analysis(df_full_path, index1_n, index2_n)
   
   Total.Results$analysis = analysis_summary
   
   #최적 path만을 따로 저장하고 Total.Results에 보관.
   Warping_Path = subset(df_full_path, Line == which(analysis_summary$Mean_of_RMSE == min(analysis_summary$Mean_of_RMSE)))
   
   Total.Results$warpingpath = Warping_Path
   
   #출력 결과를 그래프로 나타낸다.
   
   graph = ft_graph(df_full_path, analysis_summary, df_refquery, Warping_Path)
   
   #그래프들을 Total.Results에 저장한다.
   Total.Results$graph_rawdata = graph$rawdata
   Total.Results$graph_allpath = graph$allpath
   Total.Results$graph_warpingpath = graph$warpingpath
   Total.Results$graph_matching = graph$matching
   
   #Total.Results를 엑셀 파일로 저장한다.
   #ft_Save_TotalResults(Total.Results, target = 0)
   
   
   return(Total.Results)
}

main_dtw_Default_Multi = function(df_data, method=0, process = 0, reference=1, query=2){
   #method=0 : 데이터의 첫번째 열을 reference로 두번째 열을 query로 인식하여 처리한다. Default임. 
   #method=1 : 데이터를 stance phase와 swing phase로 나누어 처리한 후 합치는 방식임.
   #process=0 : 첫 번째 열과 두 번째 열만 분석 대상으로 한다. 
   #process=1 : 데이터 내에 비교 가능한 모든 열을 비교한다. 
   
   #임시로 부여한 것임.
   #reference = 1; query = 2; method = 0
   
   #엑셀 데이터를 읽어들여 data.frame을 리턴한다. 두 문장 중 하나 제거.
   # df_rawdata = ft_ReadExcel()
   df_refquery = df_data
   
   #print(method); print(process); print(reference); print(query)
   
   #분석에 필요한 열만을 df_refquery에 저장한다. reference와 query에 지정한 열만을 추출하는 과정임.
   #NA 값은 제거되지 않은 상태임.
   # if(process == 0){ # 이 부분은 중복된 것으로 파악됨. 
   #    vt_list = c(paste0("index",reference), paste0("index", query))
   #    df_refquery = subset(df_rawdata, select=c(vt_list))
   # }
   
   #Start----------------------------------------------------
   #결과 모음 리스트인 lst_Total.Results에 rawdata를 입력한다.
   Total.Results = list()
   
   #분석에 사용될 데이터에서 NA 값을 제거하기 위한 작업임.
   temp.ref = df_refquery[, 1] ; temp.query = df_refquery[, 2]
   ref.data = temp.ref[!is.na(temp.ref)]
   query.data = temp.query[!is.na(temp.query)]
   
   #NA가 제거된 자료를 벡터 형태로 Total.Results에 저장한다.
   Total.Results$index1 = data.frame("index1" = ref.data)
   Total.Results$index2 = data.frame("index2" = query.data)
   # ref_name = paste0("index", reference)
   # query_name = paste0("index", query)
   # Total.Results$index1 = data.frame(ref_name = ref.data)
   # Total.Results$index2 = data.frame(query_name = query.data)
   
   #End======================================================
   
   #cost matrix를 만들어 mt_cost에 저장한다.
   mt_cost = ft_CostMatrix(df_refquery)
   
   Total.Results$costmatrix = mt_cost
   
   #mt_cost를 보내 첫 번째 warping path를 찾는다.
   lst_warping_path = ft_FirstWarpingPath(mt_cost)
   
   #첫번째 이외의 다른 warping path를 찾아 df_full_path로 저장한다.
   df_full_path <<- ft_OtherWarpingPath(lst_warping_path, mt_cost)
   
   #RMSE를 계산하여 df_rmse에 저장한다.
   df_rmse <<- ft_RMSE(df_full_path)
   
   #위에서 구해진 df_rmse를 df_full_path와 합친다.
   df_full_path = cbind(df_full_path, df_rmse)
   
   #Start--------------------------------------------------------------------------------------------
   #Total.Results에 보관하기 위해 df_full_path의 구조를 수정한다.
   #index 값이 큰 값부터 저장된 것을 작은 값부터 저장하는 형태로 변경한다.
   df_full_path = arrange(.data = df_full_path, Line, RowIndex, ColIndex)
   
   #df_full_path$Index 값을 정렬하는 것임.
   for(i in 1:max(df_full_path$Line)){ 
      index_starter = which(df_full_path$Line == i & df_full_path$RowIndex == 1 & df_full_path$ColIndex == 1)
      index_counter = nrow(df_full_path %>% filter(Line == i))
      df_full_path$Index[index_starter:(index_starter + index_counter - 1)] = c(1:index_counter)
   }
   
   Total.Results$fullpath = df_full_path
   #End==============================================================================================
   
   #df_full_path를 분석하기 위해 수행: 분석에 사용된 자료의 길이를 계산.
   index1_n = length(ref.data)
   index2_n = length(query.data)
   
   analysis_summary = ft_analysis(df_full_path, index1_n, index2_n)
   
   Total.Results$analysis = analysis_summary
   
   #최적 path만을 따로 저장하고 Total.Results에 보관.
   Warping_Path = subset(df_full_path, Line == which(analysis_summary$Mean_of_RMSE == min(analysis_summary$Mean_of_RMSE)))
   
   Total.Results$warpingpath = Warping_Path
   
   #출력 결과를 그래프로 나타낸다.
   
   # graph = ft_graph(df_full_path, analysis_summary, df_refquery, Warping_Path)
   # 
   # #그래프들을 Total.Results에 저장한다.
   # Total.Results$graph_rawdata = graph$rawdata
   # Total.Results$graph_allpath = graph$allpath
   # Total.Results$graph_warpingpath = graph$warpingpath
   # Total.Results$graph_matching = graph$matching
   
   #Total.Results를 엑셀 파일로 저장한다.
   #ft_Save_TotalResults(Total.Results, target = 0)
   
   return(Total.Results)
}

#===디폴트인 두 열의 전체 신호 대상 분석 함수
ft_Process1 = function(method = 0, process = 0, first.index = 1, second.index = 2){
   
   if(process == 0){
      df_rawdata = ft_ReadExcel()
   }
   
   #분석할 데이터 두개 열만을 대상으로 한다.
   vt_list = c(paste0("index",first.index), paste0("index", second.index))
   df_refquery = subset(df_rawdata, select=c(vt_list))
      
   Total.Results = main_dtw(df_data=df_refquery, method = 0, process=0, reference = first.index, query = second.index)
   ft_Save_TotalResults(Total.Results, target = 0)
}

ft_Process2 = function(df_data, i, j){
   
   df_refquery = df_data
   
   Total.Results = main_dtw_Default_Multi(df_data=df_refquery, method = 0, process=0, reference = i, query = j)
   #ft_Save_Default_Multi(Total.Results, target = 0)
   return(Total.Results)
}

#===두 열을 대상으로 지지구간과 스윙구간을 나누어 분석하는 함수
ft_Process3 = function(method = 1, process = 0, first.index = 1, second.index = 2){
   
   df_rawdata = ft_ReadExcel()
   df_phase = ft_ReadExcel(filetype = 1)
   
   # Contact phase에 해당하는 데이터를 추출한다.
   temp.ref_pre = df_rawdata[as.numeric(rownames(df_rawdata)) <= df_phase[1,1], 1]
   temp.query_pre = df_rawdata[as.numeric(rownames(df_rawdata)) <= df_phase[1,2], 2]
   
   #데이터프레임을 만들기 위해 NA를 입력하여 두 열의 길이를 동일하게 만든다.
   max.len = max(length(temp.ref_pre), length(temp.query_pre))
   temp.ref_pre = c(temp.ref_pre, rep(NA, max.len - length(temp.ref_pre)))
   temp.query_pre = c(temp.query_pre, rep(NA, max.len - length(temp.query_pre)))
   sending_data_pre = data.frame("index1"=temp.ref_pre, "index2"=temp.query_pre)
   
   
   lst_PreResults = main_dtw(df_data=sending_data_pre, method = 1, process = 0, reference = first.index, query = second.index)
   lst_PreResults$warpingpath$Phase = "Pre"
   
   # Swing phase에 해당하는 데이터를 추출한다.
   temp.ref_post = df_rawdata[as.numeric(rownames(df_rawdata)) > df_phase[1,1], 1]
   temp.query_post = df_rawdata[as.numeric(rownames(df_rawdata)) > df_phase[1,2], 2]
   
   #데이터프레임을 만들기 위해 NA를 입력하여 두 열의 길이를 동일하게 만든다.
   max.len = max(length(temp.ref_post), length(temp.query_post))
   temp.ref_post = c(temp.ref_post, rep(NA, max.len - length(temp.ref_post)))
   temp.query_post = c(temp.query_post, rep(NA, max.len - length(temp.query_post)))
   sending_data_post = data.frame("index1"=temp.ref_post, "index2"=temp.query_post)
   
   lst_PostResults = main_dtw(df_data=sending_data_post, method = 1, process=0, reference = first.index, query = second.index)
   lst_PostResults$warpingpath$Phase = "Post"
   
   # Contact phase와 Swing phase를 연결하기 위해 post의 첫번째 행값을 제거한 후 결과를 결합한다.
   # lst_PostResults$warpingpath = subset(lst_PostResults$warpingpath, lst_PostResults$warpingpath$Index != 1)
   
   
   # 파일 저장을 위해 Results 내의 이름들을 변경한 후 하나의 리스트로 합친다.
   names(lst_PreResults) = c("pre_index1", "pre_index2", "pre_costmatrix", "pre_fullpath", "pre_analysis", "pre_warpingpath",
                             "pre_graph_rawdata", "pre_graph_allpath", "pre_graph_warpingpath", "pre_graph_matching")
   names(lst_PostResults) = c("post_index1", "post_index2", "post_costmatrix", "post_fullpath", "post_analysis", "post_warpingpath",
                              "post_graph_rawdata", "post_graph_allpath", "post_graph_warpingpath", "post_graph_matching")
   lst_MergeResults = c(lst_PreResults, lst_PostResults)
   
   lst_PostResults$post_warpingpath$RowIndex = lst_PostResults$post_warpingpath$RowIndex + (df_phase[1, 1])
   lst_PostResults$post_warpingpath$ColIndex = lst_PostResults$post_warpingpath$ColIndex + (df_phase[1, 2])
   
   lst_PostResults$post_warpingpath$Index = lst_PostResults$post_warpingpath$Index + (max(lst_PreResults$pre_warpingpath$Index))
   
   lst_MergeResults$MergedWarpingpath = rbind(lst_PreResults$pre_warpingpath, lst_PostResults$post_warpingpath)
   
   df_Merged_Analysis = data.frame(Phase_Name=character(), Index1_Length=numeric(), Index2_Length=numeric(), Path_Index=numeric(),
                                   Sum_of_Cost=numeric(), Cost_div_N=numeric(), Cost_div_diagonal=numeric(), 
                                   Sum_of_RMSE=numeric(), RMSE_div_N=numeric(), RMSE_div_diagonal=numeric(), Mean_of_RMSE=numeric())
   
   df_Merged_Analysis = rbind(df_Merged_Analysis, list("Pre", lst_MergeResults$pre_analysis[1, 2], lst_MergeResults$pre_analysis[1, 3], 
                                                       max(lst_MergeResults$pre_warpingpath$Index), #4. Path_Index
                                                       sum(lst_MergeResults$pre_warpingpath$Cost),  #5. Sum_of_Cost
                                                       sum(lst_MergeResults$pre_warpingpath$Cost)/
                                                          (lst_MergeResults$pre_analysis[1, 2]+lst_MergeResults$pre_analysis[1, 3]), #6. Cost_div_N
                                                       sum(lst_MergeResults$pre_warpingpath$Cost)/
                                                          sqrt(lst_MergeResults$pre_analysis[1, 2]^2 + lst_MergeResults$pre_analysis[1, 3]^2),#7. Cost_div_diagonal 
                                                       sum(lst_MergeResults$pre_warpingpath$rmse),  #8. Sum_of_RMSE
                                                       sum(lst_MergeResults$pre_warpingpath$rmse)/
                                                          (lst_MergeResults$pre_analysis[1, 2]+lst_MergeResults$pre_analysis[1, 3]), #9. RMSE_div_N
                                                       sum(lst_MergeResults$pre_warpingpath$rmse)/
                                                          sqrt(lst_MergeResults$pre_analysis[1, 2]^2 + lst_MergeResults$pre_analysis[1, 3]^2), #10. RMSE_div_diagonal
                                                       mean(lst_MergeResults$pre_warpingpath$rmse))) #11. Mean_of_RMSE
   colnames(df_Merged_Analysis) = c("Phase_Name", "Index1_Length", "Index2_Length", "Path_Index", "Sum_of_Cost", 
                                    "Cost_div_N", "Cost_div_diagonal", "Sum_of_RMSE", "RMSE_div_N", "RMSE_div_Diagonal", "Mean_of_RMSE")
   
   df_Merged_Analysis = rbind(df_Merged_Analysis, list("Post", lst_MergeResults$post_analysis[1, 2], lst_MergeResults$post_analysis[1, 3],
                                                       max(lst_MergeResults$post_warpingpath$Index), #4. Path_Index
                                                       sum(lst_MergeResults$post_warpingpath$Cost),  #5. Sum_of_Cost
                                                       sum(lst_MergeResults$post_warpingpath$Cost)/
                                                          (lst_MergeResults$post_analysis[1, 2] + lst_MergeResults$post_analysis[1, 3]),#6. Cost_div_N
                                                       sum(lst_MergeResults$post_warpingpath$Cost)/
                                                          sqrt(lst_MergeResults$post_analysis[1, 2]^2 + lst_MergeResults$post_analysis[1, 3]^2), #7. Cost_div_diagonal
                                                       sum(lst_MergeResults$post_warpingpath$rmse),  #8. Sum_of_RMSE
                                                       sum(lst_MergeResults$post_warpingpath$rmse)/
                                                          (lst_MergeResults$post_analysis[1, 2] + lst_MergeResults$post_analysis[1, 3]),#9. RMSE_div_N
                                                       sum(lst_MergeResults$post_warpingpath$rmse)/
                                                          sqrt(lst_MergeResults$post_analysis[1, 2]^2 + lst_MergeResults$post_analysis[1, 3]^2), #10. RMSE_div_diagonal
                                                       mean(lst_MergeResults$post_warpingpath$rmse))) #11. Mean_of_RMSE
   
   df_Merged_Analysis = rbind(df_Merged_Analysis, list("Merged", df_Merged_Analysis[1,2]+(df_Merged_Analysis[2, 2]), df_Merged_Analysis[1, 3] + df_Merged_Analysis[2,3],
                                                       max(lst_MergeResults$MergedWarpingpath$Index), #4. Path_Index
                                                       sum(lst_MergeResults$MergedWarpingpath$Cost),  #5. Sum_of_Cost
                                                       sum(lst_MergeResults$MergedWarpingpath$Cost)/
                                                          (df_Merged_Analysis[1,2]+df_Merged_Analysis[2, 2]+df_Merged_Analysis[1, 3]+df_Merged_Analysis[2,3]),#6. Cost_div_N
                                                       sum(lst_MergeResults$MergedWarpingpath$Cost)/
                                                          sqrt((df_Merged_Analysis[1,2]+df_Merged_Analysis[2, 2])^2 + (df_Merged_Analysis[1, 3]+df_Merged_Analysis[2,3])^2), #7. Cost_div_diagonal
                                                       sum(lst_MergeResults$MergedWarpingpath$rmse),  #8. RMSE_of_Cost
                                                       sum(lst_MergeResults$MergedWarpingpath$rmse)/
                                                          (df_Merged_Analysis[1,2]+df_Merged_Analysis[2, 2]+df_Merged_Analysis[1, 3]+df_Merged_Analysis[2,3]),#9. RMSE_div_N
                                                       sum(lst_MergeResults$MergedWarpingpath$rmse)/
                                                          sqrt((df_Merged_Analysis[1,2]+df_Merged_Analysis[2, 2])^2 + (df_Merged_Analysis[1, 3]+df_Merged_Analysis[2,3])^2), #10. RMSE_div_diagonal
                                                       mean(lst_MergeResults$MergedWarpingpath$rmse))) #11. Mean_of_RMSE
   
   lst_MergeResults$MergedAnalysis = df_Merged_Analysis
   
   #여기서부터 Merged 그래프 작성
   #Merged warping path 출력
   graph_Merged_Warpingpath = ggplot(data = lst_MergeResults$MergedWarpingpath, aes(x=ColIndex, y=RowIndex)) +
      geom_line(size=0.2) + geom_point(size = 0.2) + geom_segment(x=1, xend=df_phase[1, 2], y=1, yend=df_phase[1,1], color = "red", size=0.2) +
      geom_segment(x=df_phase[1,2]+1, xend=max(lst_MergeResults$MergedWarpingpath$ColIndex),
                   y=df_phase[1,1]+1, yend=max(lst_MergeResults$MergedWarpingpath$RowIndex), color = "red", size=0.2) 
      #theme_bw()
   
   graph_Merged_Warpingpath
   #여기서부터 Merged 결과의 matching graph 작성
   palette_rawdata = c("index1"="#000000", "index2"="#0000ff", "axis"="#797980", "red"="#ff0000")
   temp.pre.index1.length = 1:nrow(lst_MergeResults$pre_index1); temp.pre.index1.data = lst_MergeResults$pre_index1
   temp.pre.index1 = data.frame(temp.pre.index1.length, temp.pre.index1.data)
   temp.pre.index2.length = 1:nrow(lst_MergeResults$pre_index2); temp.pre.index2.data = lst_MergeResults$pre_index2
   temp.pre.index2 = data.frame(temp.pre.index2.length, temp.pre.index2.data)
   
   temp.post.index1.length = (length(temp.pre.index1.length)+1):((length(temp.pre.index1.length)+1) + (nrow(lst_MergeResults$post_index1)-1))
   temp.post.index1.data = lst_MergeResults$post_index1
   temp.post.index1 = data.frame(temp.post.index1.length, temp.post.index1.data)
   temp.post.index2.length = (length(temp.pre.index2.length)+1):((length(temp.pre.index2.length)+1) + (nrow(lst_MergeResults$post_index2)-1)); 
   temp.post.index2.data = lst_MergeResults$post_index2
   temp.post.index2 = data.frame(temp.post.index2.length, temp.post.index2.data)
   
   x_axis_length = max((length(temp.pre.index1.length)+length(temp.post.index1.length)), (length(temp.pre.index2.length)+length(temp.post.index2.length)))
   
   
   temp.ref_prevalue = temp.ref_pre[!is.na(temp.ref_pre)]
   temp.ref_postvalue = temp.ref_post[!is.na(temp.ref_post)]
   temp.query_prevalue = temp.query_pre[!is.na(temp.query_pre)]
   temp.query_postvalue = temp.query_post[!is.na(temp.query_post)]
   
   tempdata_index1 = c(temp.ref_prevalue, temp.ref_postvalue); tempdata_index2 = c(temp.query_prevalue, temp.query_postvalue)
   
   graph_Merged_Matching <- ggplot() +
      geom_point(data = temp.pre.index1, aes(x=temp.pre.index1.length, y=index1, colour = "index1"), size=0.2) +
      geom_point(data = temp.pre.index2, aes(x=temp.pre.index2.length, y=index2, colour = "index2"), size=0.2) +
      geom_point(data = temp.post.index1, aes(x=temp.post.index1.length, y=index1, colour = "index1"), size=0.2) +
      geom_point(data = temp.post.index2, aes(x=temp.post.index2.length, y=index2, colour = "index2"), size=0.2) +
      geom_hline(yintercept = 0, color = "black") +
      scale_color_manual(values = palette_rawdata) +
      xlab("Frames") +
      ylab("Values") +
      #coord_fixed(ratio = 1) 
      scale_x_continuous(breaks = seq(0, x_axis_length, 10)) +
      #scale_y_continuous(breaks = round(seq(min(rawdata_index1), max(rawdata_index1), length.out = 11), digits = 2)) +
      ggtitle("Matching Graph for Raw data") +
      theme_bw()
   #graph_Merged_Matching
   for (i in 1:length(lst_MergeResults$MergedWarpingpath$Index)){
      graph_Merged_Matching <- graph_Merged_Matching +
         annotate("segment", x = lst_MergeResults$MergedWarpingpath$ColIndex[i], xend = lst_MergeResults$MergedWarpingpath$RowIndex[i],
                  y = tempdata_index2[lst_MergeResults$MergedWarpingpath$ColIndex[i]],
                  yend = tempdata_index1[lst_MergeResults$MergedWarpingpath$RowIndex[i]], color = "red", size=0.2)
      
      #print(lst_MergeResults$MergedWarpingpath$ColIndex[i])
      #print(lst_MergeResults$MergedWarpingpath$RowIndex[i])
      #print(tempdata_index1[lst_MergeResults$MergedWarpingpath$RowIndex[i]])
   }
   
   lst_MergeResults$Merged_graph_Warpingpath = graph_Merged_Warpingpath
   lst_MergeResults$Merged_graph_Matching = graph_Merged_Matching
   
   ft_Save_phase(lst_MergeResults)

}

#=== 모든 열에 대해서 지지구간과 스윙구간으로 나누어 분석하는 함수.
ft_Process4 = function(df_data1, df_data2, i, j){
   
   df_rawdata = df_data1
   df_phase = df_data2
   
   # Contact phase에 해당하는 데이터를 추출한다.
   temp.ref_pre = df_rawdata[as.numeric(rownames(df_rawdata)) <= df_phase[1,1], 1]
   temp.query_pre = df_rawdata[as.numeric(rownames(df_rawdata)) <= df_phase[1,2], 2]
   
   #데이터프레임을 만들기 위해 NA를 입력하여 두 열의 길이를 동일하게 만든다.
   max.len = max(length(temp.ref_pre), length(temp.query_pre))
   temp.ref_pre = c(temp.ref_pre, rep(NA, max.len - length(temp.ref_pre)))
   temp.query_pre = c(temp.query_pre, rep(NA, max.len - length(temp.query_pre)))
   sending_data_pre = data.frame("index1"=temp.ref_pre, "index2"=temp.query_pre)
   
   
   lst_PreResults = main_dtw(df_data=sending_data_pre, method = 1, process = 0, reference = 1, query = 2)
   lst_PreResults$warpingpath$Phase = "Pre"
   
   # Swing phase에 해당하는 데이터를 추출한다.
   temp.ref_post = df_rawdata[as.numeric(rownames(df_rawdata)) > df_phase[1,1], 1]
   temp.query_post = df_rawdata[as.numeric(rownames(df_rawdata)) > df_phase[1,2], 2]
   
   #데이터프레임을 만들기 위해 NA를 입력하여 두 열의 길이를 동일하게 만든다.
   max.len = max(length(temp.ref_post), length(temp.query_post))
   temp.ref_post = c(temp.ref_post, rep(NA, max.len - length(temp.ref_post)))
   temp.query_post = c(temp.query_post, rep(NA, max.len - length(temp.query_post)))
   sending_data_post = data.frame("index1"=temp.ref_post, "index2"=temp.query_post)
   
   lst_PostResults = main_dtw(df_data=sending_data_post, method = 1, process=0, reference = 1, query = 2)
   lst_PostResults$warpingpath$Phase = "Post"
   
   # Contact phase와 Swing phase를 연결하기 위해 post의 첫번째 행값을 제거한 후 결과를 결합한다.
   # lst_PostResults$warpingpath = subset(lst_PostResults$warpingpath, lst_PostResults$warpingpath$Index != 1)
   
   
   # 파일 저장을 위해 Results 내의 이름들을 변경한 후 하나의 리스트로 합친다.
   names(lst_PreResults) = c("pre_index1", "pre_index2", "pre_costmatrix", "pre_fullpath", "pre_analysis", "pre_warpingpath",
                             "pre_graph_rawdata", "pre_graph_allpath", "pre_graph_warpingpath", "pre_graph_matching")
   names(lst_PostResults) = c("post_index1", "post_index2", "post_costmatrix", "post_fullpath", "post_analysis", "post_warpingpath",
                              "post_graph_rawdata", "post_graph_allpath", "post_graph_warpingpath", "post_graph_matching")
   lst_MergeResults = c(lst_PreResults, lst_PostResults)
   
   lst_PostResults$post_warpingpath$RowIndex = lst_PostResults$post_warpingpath$RowIndex + (df_phase[1, 1])
   lst_PostResults$post_warpingpath$ColIndex = lst_PostResults$post_warpingpath$ColIndex + (df_phase[1, 2])
   
   lst_PostResults$post_warpingpath$Index = lst_PostResults$post_warpingpath$Index + (max(lst_PreResults$pre_warpingpath$Index))
   
   lst_MergeResults$MergedWarpingpath = rbind(lst_PreResults$pre_warpingpath, lst_PostResults$post_warpingpath)
   
   df_Merged_Analysis = data.frame(Phase_Name=character(), Index1_Length=numeric(), Index2_Length=numeric(), Path_Index=numeric(),
                                   Sum_of_Cost=numeric(), Cost_div_N=numeric(), Cost_div_diagonal=numeric(), 
                                   Sum_of_RMSE=numeric(), RMSE_div_N=numeric(), RMSE_div_diagonal=numeric(), Mean_of_RMSE=numeric())
   
   df_Merged_Analysis = rbind(df_Merged_Analysis, list("Pre", lst_MergeResults$pre_analysis[1, 2], lst_MergeResults$pre_analysis[1, 3], 
                                                       max(lst_MergeResults$pre_warpingpath$Index), #4. Path_Index
                                                       sum(lst_MergeResults$pre_warpingpath$Cost),  #5. Sum_of_Cost
                                                       sum(lst_MergeResults$pre_warpingpath$Cost)/
                                                          (lst_MergeResults$pre_analysis[1, 2]+lst_MergeResults$pre_analysis[1, 3]), #6. Cost_div_N
                                                       sum(lst_MergeResults$pre_warpingpath$Cost)/
                                                          sqrt(lst_MergeResults$pre_analysis[1, 2]^2 + lst_MergeResults$pre_analysis[1, 3]^2),#7. Cost_div_diagonal 
                                                       sum(lst_MergeResults$pre_warpingpath$rmse),  #8. Sum_of_RMSE
                                                       sum(lst_MergeResults$pre_warpingpath$rmse)/
                                                          (lst_MergeResults$pre_analysis[1, 2]+lst_MergeResults$pre_analysis[1, 3]), #9. RMSE_div_N
                                                       sum(lst_MergeResults$pre_warpingpath$rmse)/
                                                          sqrt(lst_MergeResults$pre_analysis[1, 2]^2 + lst_MergeResults$pre_analysis[1, 3]^2), #10. RMSE_div_diagonal
                                                       mean(lst_MergeResults$pre_warpingpath$rmse))) #11. Mean_of_RMSE
   colnames(df_Merged_Analysis) = c("Phase_Name", "Index1_Length", "Index2_Length", "Path_Index", "Sum_of_Cost", 
                                    "Cost_div_N", "Cost_div_diagonal", "Sum_of_RMSE", "RMSE_div_N", "RMSE_div_Diagonal", "Mean_of_RMSE")
   
   df_Merged_Analysis = rbind(df_Merged_Analysis, list("Post", lst_MergeResults$post_analysis[1, 2], lst_MergeResults$post_analysis[1, 3],
                                                       max(lst_MergeResults$post_warpingpath$Index), #4. Path_Index
                                                       sum(lst_MergeResults$post_warpingpath$Cost),  #5. Sum_of_Cost
                                                       sum(lst_MergeResults$post_warpingpath$Cost)/
                                                          (lst_MergeResults$post_analysis[1, 2] + lst_MergeResults$post_analysis[1, 3]),#6. Cost_div_N
                                                       sum(lst_MergeResults$post_warpingpath$Cost)/
                                                          sqrt(lst_MergeResults$post_analysis[1, 2]^2 + lst_MergeResults$post_analysis[1, 3]^2), #7. Cost_div_diagonal
                                                       sum(lst_MergeResults$post_warpingpath$rmse),  #8. Sum_of_RMSE
                                                       sum(lst_MergeResults$post_warpingpath$rmse)/
                                                          (lst_MergeResults$post_analysis[1, 2] + lst_MergeResults$post_analysis[1, 3]),#9. RMSE_div_N
                                                       sum(lst_MergeResults$post_warpingpath$rmse)/
                                                          sqrt(lst_MergeResults$post_analysis[1, 2]^2 + lst_MergeResults$post_analysis[1, 3]^2), #10. RMSE_div_diagonal
                                                       mean(lst_MergeResults$post_warpingpath$rmse))) #11. Mean_of_RMSE
   
   df_Merged_Analysis = rbind(df_Merged_Analysis, list("Merged", df_Merged_Analysis[1,2]+(df_Merged_Analysis[2, 2]), df_Merged_Analysis[1, 3] + df_Merged_Analysis[2,3],
                                                       max(lst_MergeResults$MergedWarpingpath$Index), #4. Path_Index
                                                       sum(lst_MergeResults$MergedWarpingpath$Cost),  #5. Sum_of_Cost
                                                       sum(lst_MergeResults$MergedWarpingpath$Cost)/
                                                          (df_Merged_Analysis[1,2]+df_Merged_Analysis[2, 2]+df_Merged_Analysis[1, 3]+df_Merged_Analysis[2,3]),#6. Cost_div_N
                                                       sum(lst_MergeResults$MergedWarpingpath$Cost)/
                                                          sqrt((df_Merged_Analysis[1,2]+df_Merged_Analysis[2, 2])^2 + (df_Merged_Analysis[1, 3]+df_Merged_Analysis[2,3])^2), #7. Cost_div_diagonal
                                                       sum(lst_MergeResults$MergedWarpingpath$rmse),  #8. RMSE_of_Cost
                                                       sum(lst_MergeResults$MergedWarpingpath$rmse)/
                                                          (df_Merged_Analysis[1,2]+df_Merged_Analysis[2, 2]+df_Merged_Analysis[1, 3]+df_Merged_Analysis[2,3]),#9. RMSE_div_N
                                                       sum(lst_MergeResults$MergedWarpingpath$rmse)/
                                                          sqrt((df_Merged_Analysis[1,2]+df_Merged_Analysis[2, 2])^2 + (df_Merged_Analysis[1, 3]+df_Merged_Analysis[2,3])^2), #10. RMSE_div_diagonal
                                                       mean(lst_MergeResults$MergedWarpingpath$rmse))) #11. Mean_of_RMSE
   
   lst_MergeResults$MergedAnalysis = df_Merged_Analysis
   
   #여기서부터 Merged 그래프 작성
   #Merged warping path 출력
   
   # graph_Merged_Warpingpath = ggplot(data = lst_MergeResults$MergedWarpingpath, aes(x=ColIndex, y=RowIndex)) +
   #    geom_line(size=0.2) + geom_point(size = 0.2) + geom_segment(x=1, xend=df_phase[1, 2], y=1, yend=df_phase[1,1], color = "red", size=0.2) +
   #    geom_segment(x=df_phase[1,2]+1, xend=max(lst_MergeResults$MergedWarpingpath$ColIndex),
   #                 y=df_phase[1,1]+1, yend=max(lst_MergeResults$MergedWarpingpath$RowIndex), color = "red", size=0.2) 
   
   ##theme_bw()
   
   #graph_Merged_Warpingpath
   
   #여기서부터 Merged 결과의 matching graph 작성
   # palette_rawdata = c("index1"="#000000", "index2"="#0000ff", "axis"="#797980", "red"="#ff0000")
   # temp.pre.index1.length = 1:nrow(lst_MergeResults$pre_index1); temp.pre.index1.data = lst_MergeResults$pre_index1
   # temp.pre.index1 = data.frame(temp.pre.index1.length, temp.pre.index1.data)
   # temp.pre.index2.length = 1:nrow(lst_MergeResults$pre_index2); temp.pre.index2.data = lst_MergeResults$pre_index2
   # temp.pre.index2 = data.frame(temp.pre.index2.length, temp.pre.index2.data)
   # 
   # temp.post.index1.length = (length(temp.pre.index1.length)+1):((length(temp.pre.index1.length)+1) + (nrow(lst_MergeResults$post_index1)-1))
   # temp.post.index1.data = lst_MergeResults$post_index1
   # temp.post.index1 = data.frame(temp.post.index1.length, temp.post.index1.data)
   # temp.post.index2.length = (length(temp.pre.index2.length)+1):((length(temp.pre.index2.length)+1) + (nrow(lst_MergeResults$post_index2)-1)); 
   # temp.post.index2.data = lst_MergeResults$post_index2
   # temp.post.index2 = data.frame(temp.post.index2.length, temp.post.index2.data)
   # 
   # x_axis_length = max((length(temp.pre.index1.length)+length(temp.post.index1.length)), (length(temp.pre.index2.length)+length(temp.post.index2.length)))
   # 
   # 
   # temp.ref_prevalue = temp.ref_pre[!is.na(temp.ref_pre)]
   # temp.ref_postvalue = temp.ref_post[!is.na(temp.ref_post)]
   # temp.query_prevalue = temp.query_pre[!is.na(temp.query_pre)]
   # temp.query_postvalue = temp.query_post[!is.na(temp.query_post)]
   # 
   # tempdata_index1 = c(temp.ref_prevalue, temp.ref_postvalue); tempdata_index2 = c(temp.query_prevalue, temp.query_postvalue)
   # 
   # graph_Merged_Matching <- ggplot() +
   #    geom_point(data = temp.pre.index1, aes(x=temp.pre.index1.length, y=index1, colour = "index1"), size=0.2) +
   #    geom_point(data = temp.pre.index2, aes(x=temp.pre.index2.length, y=index2, colour = "index2"), size=0.2) +
   #    geom_point(data = temp.post.index1, aes(x=temp.post.index1.length, y=index1, colour = "index1"), size=0.2) +
   #    geom_point(data = temp.post.index2, aes(x=temp.post.index2.length, y=index2, colour = "index2"), size=0.2) +
   #    geom_hline(yintercept = 0, color = "black") +
   #    scale_color_manual(values = palette_rawdata) +
   #    xlab("Frames") +
   #    ylab("Values") +
   #    #coord_fixed(ratio = 1) 
   #    scale_x_continuous(breaks = seq(0, x_axis_length, 10)) +
   #    #scale_y_continuous(breaks = round(seq(min(rawdata_index1), max(rawdata_index1), length.out = 11), digits = 2)) +
   #    ggtitle("Matching Graph for Raw data") +
   #    theme_bw()
   # #graph_Merged_Matching
   # for (i in 1:length(lst_MergeResults$MergedWarpingpath$Index)){
   #    graph_Merged_Matching <- graph_Merged_Matching +
   #       annotate("segment", x = lst_MergeResults$MergedWarpingpath$ColIndex[i], xend = lst_MergeResults$MergedWarpingpath$RowIndex[i],
   #                y = tempdata_index2[lst_MergeResults$MergedWarpingpath$ColIndex[i]],
   #                yend = tempdata_index1[lst_MergeResults$MergedWarpingpath$RowIndex[i]], color = "red", size=0.2)
   #    
   #    #print(lst_MergeResults$MergedWarpingpath$ColIndex[i])
   #    #print(lst_MergeResults$MergedWarpingpath$RowIndex[i])
   #    #print(tempdata_index1[lst_MergeResults$MergedWarpingpath$RowIndex[i]])
   # }
   # 
   # lst_MergeResults$Merged_graph_Warpingpath = graph_Merged_Warpingpath
   # lst_MergeResults$Merged_graph_Matching = graph_Merged_Matching
   
   #ft_Save_Phase_Multi(lst_MergeResults)
   return(lst_MergeResults)
   
}


##===================== gDTW를 위한 Main함수==================================================
# @@@ : 이부분 나중에 활성화 할 것.
gdtw = function(method = 1){
   #method=1 : Phase 구별 없이 두 열만을 대상으로 처리함. 데이터의 첫번째 열을 reference로 두번째 열을 query로 인식하여 처리한다. Default임. 
   #method=2 : Phase 구별 없이 데이터에 있는 모든 열을 대상으로 처리함.
   #method=3 : 데이터를 stance phase와 swing phase로 나누어 처리한 후 합치는 방식임. 두 열만을 대상으로 함.
   #method=4 : 데이터를 stance phase와 swing phase로 나누어 데이터에 있는 모든 열을 대상으로 함.
   #process=0 : 첫 번째 열과 두 번째 열만 분석 대상으로 한다. 
   #process=1 : 데이터 내에 비교 가능한 모든 열을 비교한다. 
   
   #rm(list=ls())
   method = 8; process = 0 ; refer.col = 1; query.col = 2
   
   # Default mode로 phase 구분 없이 데이터의 첫번째 열을 reference로 두번째 열을 query로 인식하여 처리한다.
   if(method == 1){  
      
      ft_Process1(method = 0, process = 0, first.index = 1, second.index = 2)
   }
   
   
   #Default 형태로 phase 구분 없이 복수열의 전체 데이터를 처리함.
   if(method == 2){ 
      
      df_rawdata = ft_ReadExcel()
      column.no = ncol(df_rawdata)
      #View(df_rawdata)
      
      #>>>--엑셀 쉬트 생성------------------
      TR = createWorkbook("Total_Results")
      addWorksheet(TR, "Cost")
      addWorksheet(TR, "RMSE")
      addWorksheet(TR, "Analysis")
      #<<<<-------------------------------
      
      multi_total.result = list()
      k = 1
      excel.start.col = 1
      excel.start.row = 1
      m=1
      
      for(i in 1:(column.no-1)){
         for(j in (i+1):column.no){
            
            vt_list = c(paste0("index",i), paste0("index", j))
            df_refquery = subset(df_rawdata, select=c(vt_list))
            
            lst_Multi_Result = ft_Process2(df_data = df_refquery, i, j)
            
            #열이름을 미리 만들어 놓음.
            change.name = c(paste0("Cost_",i, "_", j), paste0("RMSE_", i, "_", j))
            # colnames(lst_Multi_Result$warpingpath[6]) = change.name[1] #열이름 Cost를 Cost_번호_번호로 변경
            # colnames(lst_Multi_Result$warpingpath[7]) = change.name[2] #열이름 rmse를 RMSE_번호_번호로 변경
            
            colnames(lst_Multi_Result$warpingpath)[6] = change.name[1]
            colnames(lst_Multi_Result$warpingpath)[7] = change.name[2]
            # print(lst_Multi_Result$warpingpath[6])
            
            # 저장을 위해 아래 두 줄을 주석 처리함.<<<<<<<<<<
            #writeData(TR, "Cost", lst_Multi_Result$warpingpath[6], colNames = TRUE, rowNames = FALSE, startCol = excel.start.col, startRow = 1)
            #writeData(TR, "RMSE", lst_Multi_Result$warpingpath[7], colNames = TRUE, rowNames = FALSE, startCol = excel.start.col, startRow = 1)
            writeData(TR, "Analysis", lst_Multi_Result[[5]], colNames = TRUE, rowNames = TRUE, startCol = "A", startRow = m)
            
            excel.start.col = excel.start.col + 1
            
            m = m + max(lst_Multi_Result$analysis$Path_No) + 1
            #print(m)
            #multi_total.result[[k]] = data.frame(lst_Multi_Result$warpingpath[6], lst_Multi_Result$warpingpath[7])
            #colnames(multi_total.result[[k]]) = change.name
            
            #multi_total.result[[k+1]] = data.frame(lst_Multi_Result$analysis)
            k = k+2
            
            #Total.Results = main_dtw(df_data=df_refquery, method = 0, process=0, reference = 1, query = 2)
         } #2th for end.
         msg = paste0(i, " 번째 완료!")
         print(msg)
      }#1th for end.
      #View(multi_total.result)
      # ft_Save_Default_Multi(multi_total.result)
      
      #폴더를 생성한다.
      folder = file.path(getwd(), "Total_Results")
      if(dir.exists(folder) == FALSE){
         dir.create(folder)
      }
      
      #파일을 생성한다.
      filecounter = 1
      while(filecounter > 0){
         fname = paste0(folder, "/Total_Results_Multi", sprintf("%03d", filecounter), ".XLSX")
         if(file.exists(fname) == FALSE){
            saveWorkbook(TR, file = fname)
            #file.create(fname)
            break
         }
         filecounter = filecounter + 1
      }
      
      print("파일 생성 완료됨.")
      
   }#if end.
   
   
   #데이터를 stance phase와 swing phase로 나누어 처리한 후 합치는 방식임. 1열과 2열만을 대상으로 한다.
   if(method == 3){ 
      
      ft_Process3(method = 1, process = 0, first.index = 1, second.index = 2)

   } #method=1 end.
   
   
   #데이터를 stance phase와 swing phase로 나누어 복수열 전체 데이터를 처리함.
   if(method == 4){ 
      
      df_rawdata = ft_ReadExcel()
      df_phase = ft_ReadExcel(filetype = 1)
      
      column.no = ncol(df_rawdata)
      
      #>>>--엑셀 쉬트 생성------------------
      TR = createWorkbook("Total_Results")
      addWorksheet(TR, "Cost")
      addWorksheet(TR, "RMSE")
      addWorksheet(TR, "Analysis")
      #<<<<-------------------------------
      
      multi_total.result = list()
      k = 1
      excel.start.col = 1
      excel.start.row = 1
      m=1
      
      for(i in 1:(column.no-1)){
         for(j in (i+1):column.no){
            
            vt_list = c(paste0("index",i), paste0("index", j))
            df_refquery = subset(df_rawdata, select=c(vt_list))
            df_Sel.Phase = subset(df_phase, select=c(vt_list))
            
            lst_Multi_Result = ft_Process4(df_data1 = df_refquery, df_data2 = df_Sel.Phase, i, j)
            
            #열이름을 미리 만들어 놓음.
            change.name = c(paste0("Cost_",i, "_", j), paste0("RMSE_", i, "_", j))
            
            colnames(lst_Multi_Result$MergedWarpingpath)[6] = change.name[1]
            colnames(lst_Multi_Result$MergedWarpingpath)[7] = change.name[2]
            #print(lst_Multi_Result$warpingpath[6])
            writeData(TR, "Cost", lst_Multi_Result$MergedWarpingpath[6], colNames = TRUE, rowNames = FALSE, startCol = excel.start.col, startRow = 1)
            writeData(TR, "RMSE", lst_Multi_Result$MergedWarpingpath[7], colNames = TRUE, rowNames = FALSE, startCol = excel.start.col, startRow = 1)
            writeData(TR, "Analysis", lst_Multi_Result$MergedAnalysis, colNames = TRUE, rowNames = TRUE, startCol = "A", startRow = m)
            
            excel.start.col = excel.start.col + 1
            m = m + 5
            
            # multi_total.result[[k]] = data.frame(lst_Multi_Result$MergedWarpingpath[6], lst_Multi_Result$MergedWarpingpath[7])
            # colnames(multi_total.result[[k]]) = change.name
            
            # multi_total.result[[k+1]] = data.frame(lst_Multi_Result$MergedAnalysis)
            k = k+2
            
            #Total.Results = main_dtw(df_data=df_refquery, method = 0, process=0, reference = 1, query = 2)
         } #2th for end.
         print("첫번째 for 완료")
      }#1th for end.
      #View(multi_total.result)
      # ft_Save_Phase_Multi(multi_total.result)
      
      folder = file.path(getwd(), "Total_Results")
      if(dir.exists(folder) == FALSE){
         dir.create(folder)
      }
      
      #파일을 생성한다.
      filecounter = 1
      while(filecounter > 0){
         fname = paste0(folder, "/Total_Results_Phase_Multi", sprintf("%03d", filecounter), ".XLSX")
         if(file.exists(fname) == FALSE){
            saveWorkbook(TR, file = fname)
            #file.create(fname)
            break
         }
         filecounter = filecounter + 1
      }
      
      print("파일 생성 완료됨.")
   }#if end.
   
   ## 임시용임. 수행한 후 제거할 것.
   ## 피험자별 분석을 위해 만든 코드임. Within subjects 자료 생성용임.
   if(method == 5){ 
      
      df_rawdata = ft_ReadExcel()
      column.no = ncol(df_rawdata)
      #View(df_rawdata)
      
      #아래 162 값을 성별에 맞게 입력하면 됨.
      sbj = seq(1, 168, 6)
      sbj
      for (yy in sbj) {
         yyy = yy+5
         df_rawdata_temp = df_rawdata[yy:yyy]#6개열씩 자료를 이동
         if(yy == 7) {
            print(colnames(df_rawdata_temp))
         }
         colnames(df_rawdata_temp) = c('index1', 'index2', 'index3', 'index4', 'index5', 'index6')
         column.no = ncol(df_rawdata_temp)
         
         #>>>--엑셀 쉬트 생성------------------
         TR = createWorkbook("Total_Results")
         addWorksheet(TR, "Cost")
         addWorksheet(TR, "RMSE")
         addWorksheet(TR, "Analysis")
         #<<<<-------------------------------
         
         multi_total.result = list()
         k = 1
         excel.start.col = 1
         excel.start.row = 1
         m=1 #엑셀파일의 시작열 지정

      for(i in 1:(column.no-1)){
         for(j in (i+1):column.no){
            
            vt_list = c(paste0("index",i), paste0("index", j)) #추출할 이름을 미리 설정함.
            df_refquery = subset(df_rawdata_temp, select=c(vt_list)) #위에서 설정한 이름에 해당하는 자료를 별도로 저장.
            
            lst_Multi_Result = ft_Process2(df_data = df_refquery, i, j) #DTW를 실행하여 저장.
            
            #열이름을 미리 만들어 놓음.
            change.name = c(paste0("Cost_",i, "_", j), paste0("RMSE_", i, "_", j))
            # colnames(lst_Multi_Result$warpingpath[6]) = change.name[1] #열이름 Cost를 Cost_번호_번호로 변경
            # colnames(lst_Multi_Result$warpingpath[7]) = change.name[2] #열이름 rmse를 RMSE_번호_번호로 변경
            
            colnames(lst_Multi_Result$warpingpath)[6] = change.name[1]
            colnames(lst_Multi_Result$warpingpath)[7] = change.name[2]
            # print(lst_Multi_Result$warpingpath[6])
            
            # 저장을 위해 아래 두 줄을 주석 처리함.<<<<<<<<<<
            #writeData(TR, "Cost", lst_Multi_Result$warpingpath[6], colNames = TRUE, rowNames = FALSE, startCol = excel.start.col, startRow = 1)
            #writeData(TR, "RMSE", lst_Multi_Result$warpingpath[7], colNames = TRUE, rowNames = FALSE, startCol = excel.start.col, startRow = 1)
            writeData(TR, "Analysis", lst_Multi_Result[[5]], colNames = TRUE, rowNames = TRUE, startCol = "A", startRow = m)
            
            excel.start.col = excel.start.col + 1
            
            m = m + max(lst_Multi_Result$analysis$Path_No) + 1
            #print(m)
            #multi_total.result[[k]] = data.frame(lst_Multi_Result$warpingpath[6], lst_Multi_Result$warpingpath[7])
            #colnames(multi_total.result[[k]]) = change.name
            
            #multi_total.result[[k+1]] = data.frame(lst_Multi_Result$analysis)
            k = k+2
            
            #Total.Results = main_dtw(df_data=df_refquery, method = 0, process=0, reference = 1, query = 2)
         } #2th for end.
         msg = paste0(i, " 번째 완료!")
         print(msg)
      }#1th for end.
      #View(multi_total.result)
      # ft_Save_Default_Multi(multi_total.result)
      
      #폴더를 생성한다.
      folder = file.path(getwd(), "Total_Results")
      if(dir.exists(folder) == FALSE){
         dir.create(folder)
      }
      
      #파일을 생성한다.
      filecounter = 1
      while(filecounter > 0){
         fname = paste0(folder, "/Total_Results_Multi", sprintf("%03d", filecounter), ".XLSX")
         if(file.exists(fname) == FALSE){
            saveWorkbook(TR, file = fname)
            #file.create(fname)
            break
         }
         filecounter = filecounter + 1
      }
      
      print("파일 생성 완료됨.")
      
      } #for end
      
   }#if end.
   
   ##임시용임. 
   ##피험자의 모든 왼발, 오른발을 성별별로 처리함.
   if(method == 6){ 
      
      df_rawdata = ft_ReadExcel()
      column.no = ncol(df_rawdata)
      #View(df_rawdata)
      
      df_rawdata_temp = df_rawdata
      variable_counter = seq(1, ncol(df_rawdata_temp))
      
      #데이터프레임의 모든 열에 이름을 지정함.
      for (i in variable_counter) {
         colnames(df_rawdata_temp)[i] = paste0("col_",i)
      }
      
      #>>>--엑셀 쉬트 생성------------------
      TR = createWorkbook("Total_Results")
      addWorksheet(TR, "Cost")
      addWorksheet(TR, "RMSE")
      addWorksheet(TR, "Analysis")
      #<<<<-------------------------------
      multi_total.result = list()
      excel.start.col = 1
      excel.start.row = 1
      m=1 #엑셀파일의 시작열 지정
         
      #아래 162(variable_counter) 값을 성별에 맞게 입력하면 됨.
      sbj_counter = seq(1, 28) #피험자 숫자.
      sbj = seq(1, column.no - 6, 6)
      colcounter = 1:column.no-6
      #colcounter = data.frame(matrix(1:162, nrow = 6))
      
      for (i in sbj) {#for1 # 1, 7, 13, ....
         jcounter = seq(i, i+5) 
         
         for (j in jcounter) {#for2
            kcounter = seq(max(jcounter)+1, column.no)
            
            for (k in kcounter) {#for3
               #print(j)
               #print(k)
               msgg = paste0(j, " _ ", k)
               print(msgg)
               df_refquery = data.frame(df_rawdata_temp[j], df_rawdata_temp[k])
               
               lst_Multi_Result = ft_Process2(df_data = df_refquery, i, j) #DTW를 실행하여 저장.
               
               change.name = c(paste0("Cost_",j, "_", k), paste0("RMSE_", j, "_", k))
               colnames(lst_Multi_Result$warpingpath)[6] = change.name[1]
               colnames(lst_Multi_Result$warpingpath)[7] = change.name[2]
               
               lst_Multi_Result$analysis = arrange(lst_Multi_Result$analysis, Path_Index, Sum_of_Cost, sum_of_RMSE, Mean_of_RMSE)
               
               writeData(TR, "Analysis", lst_Multi_Result[[5]], colNames = TRUE, rowNames = TRUE, startCol = "A", startRow = m)
               writeData(TR, "Analysis", change.name[2], colNames = FALSE, rowNames = FALSE, startCol = "N", startRow = m+1)
               excel.start.col = excel.start.col + 1
               m = m + max(lst_Multi_Result$analysis$Path_No) + 1
            }#for3
            msg = paste0(j, " 번째 완료!")
            print(msg)
         }#for2
         
      }#for1
      
      #폴더를 생성한다.
      folder = file.path(getwd(), "Total_Results")
      if(dir.exists(folder) == FALSE){
         dir.create(folder)
      }
      
      #파일을 생성한다.
      filecounter = 1
      while(filecounter > 0){
         fname = paste0(folder, "/Total_Results_Multi", sprintf("%03d", filecounter), ".XLSX")
         if(file.exists(fname) == FALSE){
            saveWorkbook(TR, file = fname)
            #file.create(fname)
            break
         }
         filecounter = filecounter + 1
      }
      
      print("파일 생성 완료됨.")
      
   } #if end
   
   ##임시용임. 피험자의 왼발vs.오른발을 비교하는 루틴임.
   if(method == 7){
      
      df_rawdata = ft_ReadExcel()
      column.no = ncol(df_rawdata)
      #View(df_rawdata)
      
      df_rawdata_temp = df_rawdata
      variable_counter = seq(1, ncol(df_rawdata_temp))
      
      #데이터프레임의 모든 열에 이름을 지정함.
      for (i in variable_counter) {
         colnames(df_rawdata_temp)[i] = paste0("col_",i)
      }
      
      #>>>--엑셀 쉬트 생성------------------
      TR = createWorkbook("Total_Results")
      addWorksheet(TR, "Cost")
      addWorksheet(TR, "RMSE")
      addWorksheet(TR, "Analysis")
      #<<<<-------------------------------
      multi_total.result = list()
      excel.start.col = 1
      excel.start.row = 1
      m=1 #엑셀파일의 시작열 지정
      
      #아래 162(variable_counter) 값을 성별에 맞게 입력하면 됨.
      # sbj_counter = seq(1, 28) #피험자 숫자.
      # sbj = seq(1, column.no - 6, 6)
      colcounter = seq(1, (column.no/2))
      #colcounter = data.frame(matrix(1:168, nrow = 6))
      
      for (i in colcounter) {#1....162까지.
         
         #print(j)
         #print(k)
         msgg = paste0(i, " _ ", i+168)
         print(msgg)
         df_refquery = data.frame(df_rawdata_temp[i], df_rawdata_temp[i+168])
               
         lst_Multi_Result = ft_Process2(df_data = df_refquery, 1, 2) #DTW를 실행하여 저장.
               
         change.name = c(paste0("Cost_",i , "_", i+168), paste0("RMSE_", i, "_", i+168))
         colnames(lst_Multi_Result$warpingpath)[6] = change.name[1]
         colnames(lst_Multi_Result$warpingpath)[7] = change.name[2]
               
         lst_Multi_Result$analysis = arrange(lst_Multi_Result$analysis, Path_Index, Sum_of_Cost, sum_of_RMSE, Mean_of_RMSE)
               
         writeData(TR, "Analysis", lst_Multi_Result[[5]], colNames = TRUE, rowNames = TRUE, startCol = "A", startRow = m)
         writeData(TR, "Analysis", change.name[2], colNames = FALSE, rowNames = FALSE, startCol = "N", startRow = m+1)
         excel.start.col = excel.start.col + 1
         m = m + max(lst_Multi_Result$analysis$Path_No) + 1
         
         if(i%%6 == 0){
            msg = paste0(i%/%6, " 번째 완료!")
            print(msg)
         }
   
      }#for1
      
      #폴더를 생성한다.
      folder = file.path(getwd(), "Total_Results")
      if(dir.exists(folder) == FALSE){
         dir.create(folder)
      }
      
      #파일을 생성한다.
      filecounter = 1
      while(filecounter > 0){
         fname = paste0(folder, "/Total_Results_Multi", sprintf("%03d", filecounter), ".XLSX")
         if(file.exists(fname) == FALSE){
            saveWorkbook(TR, file = fname)
            #file.create(fname)
            break
         }
         filecounter = filecounter + 1
      }
      
      print("파일 생성 완료됨.")
   }#if end.
   
   ##임시용임. 비정상 보행 분석용임.
   if(method == 8){
      
      df_rawdata = ft_ReadExcel()
      column.no = ncol(df_rawdata)
      #View(df_rawdata)
      
      df_rawdata_temp = df_rawdata
      variable_counter = seq(1, ncol(df_rawdata_temp))
      
      #데이터프레임의 모든 열에 이름을 지정함.
      for (i in variable_counter) {
         colnames(df_rawdata_temp)[i] = paste0("col_",i)
      }
      
      #>>>--엑셀 쉬트 생성------------------
      TR = createWorkbook("Total_Results")
      addWorksheet(TR, "Cost")
      addWorksheet(TR, "RMSE")
      addWorksheet(TR, "Analysis")
      #<<<<-------------------------------
      multi_total.result = list()
      excel.start.col = 1
      excel.start.row = 1
      m=1 #엑셀파일의 시작열 지정
      
      #첫 번째 for문을 위해 전반부 대상을 계산하여 넣음.
      colcounter = seq(1, (column.no/2))
      
      for (i in colcounter) {#여기서는 1...40까지가 됨.
         
         msgg = paste0(i, " _ ", i+max(colcounter)) #첫 번째는 1_41를 출력함.
         print(msgg)
         
         df_refquery = data.frame(df_rawdata_temp[i], df_rawdata_temp[i+max(colcounter)])
         
         lst_Multi_Result = ft_Process2(df_data = df_refquery, 1, 2) #DTW를 실행하여 저장.
         
         change.name = c(paste0("Cost_",i , "_", i+max(colcounter)), paste0("RMSE_", i, "_", i+max(colcounter)))
         colnames(lst_Multi_Result$warpingpath)[6] = change.name[1]
         colnames(lst_Multi_Result$warpingpath)[7] = change.name[2]
         
         lst_Multi_Result$analysis = arrange(lst_Multi_Result$analysis, Path_Index, Sum_of_Cost, sum_of_RMSE, Mean_of_RMSE)
         
         writeData(TR, "Analysis", lst_Multi_Result[[5]], colNames = TRUE, rowNames = TRUE, startCol = "A", startRow = m)
         writeData(TR, "Analysis", change.name[2], colNames = FALSE, rowNames = FALSE, startCol = "N", startRow = m+1)
         excel.start.col = excel.start.col + 1
         m = m + max(lst_Multi_Result$analysis$Path_No) + 1
         
         if(i%%4 == 0){
            msg = paste0(i%/%4, " 번째 완료!")
            print(msg)
         }
         
      }#for1
      
      #폴더를 생성한다.
      folder = file.path(getwd(), "Total_Results")
      if(dir.exists(folder) == FALSE){
         dir.create(folder)
      }
      
      #파일을 생성한다.
      filecounter = 1
      while(filecounter > 0){
         fname = paste0(folder, "/Total_Results_Multi", sprintf("%03d", filecounter), ".XLSX")
         if(file.exists(fname) == FALSE){
            saveWorkbook(TR, file = fname)
            #file.create(fname)
            break
         }
         filecounter = filecounter + 1
      }
      
      print("파일 생성 완료됨.")
   }#if end.
      
   
   
}#function end.

rm(lst_Multi_Result)


