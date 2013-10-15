use strict;
use Win32::OLE qw(in with);
use Win32::OLE::Const 'Microsoft Excel';


#$Win32::OLE::Warn = 3;　                                 # die on errors...
# get already active Excel application or open new

my $Excel = Win32::OLE->new("Excel.Application" , sub { $_[0]->Quit } ) 
    or die Win32::OLE::LastError;
$Excel->{Visible} = 1;

my  $file  =  'D:\little desk\perl\test2.xls' ;
my  $value  =  0 ;
#宣告一個要運算的變數
my  $book  =  $Excel -> Workbooks -> Open (  $file  );
#打開檔案 並讀取分頁(worksheet)

foreach  my  $Sheet  ( in  $book -> { Worksheets }) 
{
    my  $sheetName  =  $Sheet -> { Name };  
    print  "which worklist: $sheetName\n" ;  
    #印出分頁的名字
	#如果有多個選單 也可以多份運算
 
 	my  $minRow  =  2 ;
 	my  $maxRow  =  $Sheet -> UsedRange -> Rows -> Count ; 
	#得到row最大直
    printf ( "how many cols: %d,\n" , $maxRow );
	#印出共有幾層col
    foreach  my  $row  (  $minRow  ..  $maxRow  ){  
        	
            my  $cell_value  =  $Sheet -> Cells ( $row , 1 ) -> { Value };
			#cell抓取第幾行第幾列 sheet是xls表單
			#先直在橫 row是直 col是橫
			$value = $value+$cell_value ;
            next  unless  defined  $cell_value ;  
             
        }  
         print  "how many var total: \n" ,  $value ;
    }  


$book -> Close ();
$Excel -> Quit ();