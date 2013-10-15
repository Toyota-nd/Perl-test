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
	
	
	#undef=null 沒值
	
	my $cell_value  =  $Sheet -> Cells ( 2 , 1 ) -> { Value };
	#先抓第一筆資料
	my $now = $cell_value;
	my $val = 0;
	
    foreach  my  $row  (  $minRow  ..  $maxRow  ){  
	
            $cell_value  =  $Sheet -> Cells ( $row , 1 ) -> { Value };
			#cell抓取第幾行第幾列 sheet是xls表單
			#先直在橫 row是直 col是橫
            next  unless  defined  $cell_value ;  
			
			
			#判斷式  如果下個欄位的值不等於目前欄位
            if ($cell_value ne $now) {
			print  "$now  "."have "."$val \n";
			#送出舊值
			$now = $cell_value;
			$val = 0;
			#帶入新值
			} 
			
			$val = $val +1;
			
        }  
         
  


$book -> Close ();
$Excel -> Quit ();