use strict;
use Win32::OLE qw(in with);
use Win32::OLE::Const 'Microsoft Excel';


#$Win32::OLE::Warn = 3;　                                 # die on errors...
# get already active Excel application or open new

my $Excel = Win32::OLE->new("Excel.Application" , sub { $_[0]->Quit } ) 
    or die Win32::OLE::LastError;
$Excel->{Visible} = 1;

my  $file  =  'D:\little desk\perl\test3.xls' ;
my  $value  =  0 ;
#宣告一個要運算的變數
my  $book  =  $Excel -> Workbooks -> Open (  $file  );
#打開檔案 並讀取分頁(worksheet)




my $Sheet = $book->Worksheets(1);
#選擇哪個分頁
 	my  $minRow  =  2 ;
 	my  $maxRow  =  $Sheet -> UsedRange -> Rows -> Count ; 
	#得到row最大直
    printf ( "how many cols: %d,\n" , $maxRow );
	#印出共有幾層col
	print  "all Department in nd: \n"; 
	
	
	my $now = "null";
    foreach  my  $row  (  $minRow  ..  $maxRow  ){  
        	
            my  $cell_value  =  $Sheet -> Cells ( $row , 1 ) -> { Value };
			#cell抓取第幾行第幾列 sheet是xls表單
			#先直在橫 row是直 col是橫
            next  unless  defined  $cell_value ;  
         
     	    if ($cell_value ne $now) {
			$now = $cell_value;			
			print  "$now\n";
			#如有遇到新值才輸出 避免重複
			} 
        }  
         
  


$book -> Close ();
$Excel -> Quit ();