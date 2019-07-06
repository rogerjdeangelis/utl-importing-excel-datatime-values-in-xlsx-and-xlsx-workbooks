Importing excel datetime values in xlsx and xlsx workbooks                                                                          
                                                                                                                                    
  Although SAS can export and import excel dates there appears to be a bug when dealing with excel datetime values.                 
                                                                                                                                    
  The solution involves SAS passthru to excel                                                                                       
                                                                                                                                    
     Two Solutions (both versions of excel workbooks)                                                                               
                                                                                                                                    
           a. XLS  Workbook                                                                                                         
           b. XLSX Workbook                                                                                                         
                                                                                                                                    
                                                                                                                                    
                                                                                                                                    
other excel repos                                                                                                                   
https://tinyurl.com/y3p2pqcs                                                                                                        
https://github.com/rogerjdeangelis?utf8=%E2%9C%93&tab=repositories&q=excel+in%3Aname&type=&language=                                
                                                                                                                                    
https://tinyurl.com/y2guxahe                                                                                                        
https://communities.sas.com/t5/SAS-Programming/Import-Time-Values-from-XLS-File/m-p/571003                                          
                                                                                                                                    
*_                   _                                                                                                              
(_)_ __  _ __  _   _| |_                                                                                                            
| | '_ \| '_ \| | | | __|                                                                                                           
| | | | | |_) | |_| | |_                                                                                                            
|_|_| |_| .__/ \__,_|\__|                                                                                                           
        |_|                                                                                                                         
;                                                                                                                                   
                                                                                                                                    
* create an xls and xlsx workbook with exported SAS datetimes;                                                                      
                                                                                                                                    
%utlfkil(d:/xls/xlsxFmt.xlsx);                                                                                                      
%utlfkil(d:/xls/xlsFmt.xls);                                                                                                        
                                                                                                                                    
libname xlsxFmt "d:/xls/xlsxFmt.xlsx";                                                                                              
libname xlsFmt " d:/xls/xlsFmt.xls";                                                                                                
                                                                                                                                    
data xlsxfmt.havTymxlsx  xlsFmt.havTymxls;                                                                                          
  format datTyms datetime23.;                                                                                                       
  datTym=int(datetime());                                                                                                           
  put datTym datetime23.;                                                                                                           
  do datTyms=datTym to datTym+3600*5 by 3600;                                                                                       
      output;                                                                                                                       
  end;                                                                                                                              
  drop datTym;                                                                                                                      
run;quit;                                                                                                                           
libname xlsxFmt clear;                                                                                                              
libname xlsFmt  clear;                                                                                                              
                                                                                                                                    
run;quit;                                                                                                                           
                                                                                                                                    
                                                                                                                                    
* TWO WORKBOOKS                                                                                                                     
                                                                                                                                    
d:/xls/xlsxFmt.xlsx                                                                                                                 
d:/xls/class.xlsx                                                                                                                   
                                                                                                                                    
 Note excel displays the datetimes as dates but                                                                                     
the unelying number id datetime.                                                                                                    
                                                                                                                                    
                                                                                                                                    
    +------------+                                                                                                                  
    |    A       |                                                                                                                  
    +------------+                                                                                                                  
 1  |  datTyms   |                                                                                                                  
    +------------+                                                                                                                  
 2  |  7/6/2019  |                                                                                                                  
    +------------+                                                                                                                  
 3  |  7/6/2019  |                                                                                                                  
    +------------+                                                                                                                  
 4  |  7/6/2019  |                                                                                                                  
    +------------+                                                                                                                  
 5  |  7/6/2019  |                                                                                                                  
    +------------+                                                                                                                  
 6  |  7/6/2019  |                                                                                                                  
    +------------+                                                                                                                  
 7  |  7/6/2019  |                                                                                                                  
    +------------+                                                                                                                  
                                                                                                                                    
    [HAVTYMXLS]                                                                                                                     
                                                                                                                                    
    [HAVTYMXLSX]                                                                                                                    
                                                                                                                                    
*            _               _                                                                                                      
  ___  _   _| |_ _ __  _   _| |_                                                                                                    
 / _ \| | | | __| '_ \| | | | __|                                                                                                   
| (_) | |_| | |_| |_) | |_| | |_                                                                                                    
 \___/ \__,_|\__| .__/ \__,_|\__|                                                                                                   
                |_|                                                                                                                 
;                                                                                                                                   
                                                                                                                                    
TWO TABLES                                                                                                                          
                                                                                                                                    
WORK.DATESXls total obs=6                                                                                                           
                                                                                                                                    
WORK.DATESXlsx total obs=6                                                                                                          
                                                                                                                                    
Note date is character so youi have to covert to SAS dataetime                                                                      
                                                                                                                                    
Obs         CHRDTETYM                                                                                                               
                                                                                                                                    
 1     07-06-2019 10:47:36                                                                                                          
 2     07-06-2019 11:47:36                                                                                                          
 3     07-06-2019 12:47:36                                                                                                          
 4     07-06-2019 13:47:36                                                                                                          
 5     07-06-2019 14:47:36                                                                                                          
 6     07-06-2019 15:47:36                                                                                                          
                                                                                                                                    
                                                                                                                                    
*          _       _   _                                                                                                            
 ___  ___ | |_   _| |_(_) ___  _ __  ___                                                                                            
/ __|/ _ \| | | | | __| |/ _ \| '_ \/ __|                                                                                           
\__ \ (_) | | |_| | |_| | (_) | | | \__ \                                                                                           
|___/\___/|_|\__,_|\__|_|\___/|_| |_|___/                                                                                           
                                                                                                                                    
;                                                                                                                                   
                                                                                                                                    
*               _                                                                                                                   
  __ _    __  _| |___                                                                                                               
 / _` |   \ \/ / / __|                                                                                                              
| (_| |_   >  <| \__ \                                                                                                              
 \__,_(_) /_/\_\_|___/                                                                                                              
                                                                                                                                    
;                                                                                                                                   
                                                                                                                                    
                                                                                                                                    
proc sql;                                                                                                                           
   connect to excel (Path="d:\xls\xlsFmt.xls" mixed=yes);                                                                           
   create                                                                                                                           
       table datesXls as                                                                                                            
   select                                                                                                                           
      chrdteTym                                                                                                                     
   from                                                                                                                             
      connection to Excel                                                                                                           
       (                                                                                                                            
        Select                                                                                                                      
            Format(datTyms,"mm-dd-yyyy hh:nn:ss") as chrDteTym                                                                      
        from                                                                                                                        
             [havTymXls$]                                                                                                           
       );                                                                                                                           
;quit;                                                                                                                              
                                                                                                                                    
*_             _                                                                                                                    
| |__    __  _| |_____  __                                                                                                          
| '_ \   \ \/ / / __\ \/ /                                                                                                          
| |_) |   >  <| \__ \>  <                                                                                                           
|_.__(_) /_/\_\_|___/_/\_\                                                                                                          
                                                                                                                                    
;                                                                                                                                   
                                                                                                                                    
proc sql;                                                                                                                           
   connect to excel (Path="d:\xls\xlsxFmt.xlsx" mixed=yes);                                                                         
   create                                                                                                                           
       table datesXlsx as                                                                                                           
   select                                                                                                                           
      chrdteTym                                                                                                                     
   from                                                                                                                             
      connection to Excel                                                                                                           
       (                                                                                                                            
        Select                                                                                                                      
            Format(datTyms,"mm-dd-yyyy hh:nn:ss") as chrDteTym                                                                      
        from                                                                                                                        
             [havTymXlsx$]                                                                                                          
       );                                                                                                                           
;quit;                                                                                                                              
                                                                                                                                    
                                                                                                                                    
                                                                                                                                    
