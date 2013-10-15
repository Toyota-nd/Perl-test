#!/usr/bin/perl
use strict;
use warnings;
use Win32::OLE; 
use Win32::OLE::Const 'Microsoft Excel';  # brin in Excel constants
use FindBin;



print 'enter filename which inset: ';
chomp(my $pull=<STDIN>);



my $Excel = Win32::OLE->new("Excel.Application" , sub { $_[0]->Quit } ) 
    or die Win32::OLE::LastError;
$Excel->{Visible} = 1;

my $path = $FindBin::Bin;
#得到目前路徑

my $Book = $Excel->Workbooks->Open("$path".'/'."$pull") 
    or die Win32::OLE::LastError;
	
	
print 'enter filename which output: ';
chomp(my $push=<STDIN>);


my $Sheet = $Book->Worksheets(1);
$Excel->{DisplayAlerts} = 0;  # avoid being prompted
$Book->SaveAs("$path".'/'."$push")  
or die Win32::OLE::LastError;



#ole正常啟用 可以讀寫 只是純粹xls要轉txt要另外編碼 誤會
#證明>> ole可以讀取某xls資料 這沒問題
#因為也存了同樣內容的另外一個xls
   
$Book->Close();
$Excel->Quit;