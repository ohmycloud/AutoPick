use strict;  
use warnings;
use 5.010;
use Mojo::UserAgent;
use Mojo::UserAgent::CookieJar;
use Mojo::UserAgent::Proxy;
use Win32::API;
use MyExcelFormatter;
use Encode;
use AnyEvent;
use Excel::Writer::XLSX;

sub H{
my $text = shift;
return  decode('utf8',$text);  # 进行转码
}

sub getTime
{
    my $time = shift || time();
    my ($sec,$min,$hour,$mday,$mon,$year,$wday,$yday,$isdst) = localtime($time);

    $year += 1900;
    $mon ++;

    $min  = '0'.$min  if length($min)  < 2;
    $sec  = '0'.$sec  if length($sec)  < 2;
    $mon  = '0'.$mon  if length($mon)  < 2;
    $mday = '0'.$mday if length($mday) < 2;
    $hour = '0'.$hour if length($hour) < 2;
    
    my $weekday = ('Sun','Mon','Tue','Wed','Thu','Fri','Sat')[$wday];

    return { 'second' => $sec,
             'minute' => $min,
             'hour'   => $hour,
             'day'    => $mday,
             'month'  => $mon,
             'year'   => $year,
             'weekNo' => $wday,
             'wday'   => $weekday,
             'yday'   => $yday,
             'date'   => "$year-$mon-$mday"
          };
}


my $cv = AnyEvent->condvar;
my $count = 0;
my $w; $w = AnyEvent->timer(
        after       => 1,
        interval => 300, # 每隔 5 分钟检查一次
        cb => sub {
		     my $date = getTime();  
             my $today = $date->{date};                         # 获取xxxx-xx-xx这样的日期
             my $hour  = $date->{hour};
             my $minute =$date->{minute};                  
             #my $month = $date->{month};                        # 获取月
             #my $day = $date->{day};                            # 获取日
             #my $year = $date->{year};                          # 获取年
             #my $weekday = $date->{wday};                       # 获取星期
             my $yesterday = getTime(time() - 86400)->{date};   # 获取昨天的日期，也可以用 86400*N，获取N天前的日期

             $count++;
			 say " -> check the $count times at $hour:$minute";
       
my $ua   = Mojo::UserAgent->new(timeout => 600);
my $response = $ua->get('http://118.144.74.39:9198/commons/image.jsp');
my $login_url = 'http://118.144.74.39:9198/login.do';
my $post_url='http://118.144.74.39:9198/report/listReport.do';

 if ($response->success) {
# $response->res->content->asset->move_to('verifycode.BMP');
# my $dll={};
# my $D='AdvOcr.dll';
# $dll->{OcrInit} = Win32::API->new($D, 'OcrInit',[],'N') || die " Can't open the dll file $D $!";

# $dll->{OCR_C} = Win32::API->new($D, 'OCR_C',['P','P'],'P');
# if($dll->{OcrInit}->Call()){
# my $ocr_txt=$dll->{OCR_C}->Call('163_esales','verifycode.BMP');
# print " -> 验证码:$ocr_txt\n";

$response->res->content->asset->move_to('verifycode.BMP');

system("tesseract.exe verifycode.BMP verify 1>nul 2>nul");
open my $verify,"<","verify.txt" or die;
chomp(my $verifycode = <$verify>);
close $verify;
	

my $login_form= {
'anchor'       =>	'#/navigation/109/menugroup/blank/menu/dmsProdMappingListNe',
'rand'         =>   "$verifycode",
'userAccount'  =>	'winc_sxw',
'userPassword' =>   '000000',
};
my $tx = $ua->post( $login_url => form => $login_form );
if ( $tx->success ) {
    if ($tx->res->body !~ /验证码输入错误/) {
my $post_form={
'_RES_ID_' => '254',
'activeStoreSource' => 'OT',
# 'colIds' => '18,0,1,2,3,5,6,7,21,35,32,33,29,30,31,8,22,27,28,9,10,11,12,13,25,26,14,15,16,17,20,24,34,23,19',
'colIds' => '0,1,2,3,5,7,8,9,12,15,26',
'createTime1' => "$today",
'createTime2' => "$today",
'ec_i' => 'server_dmsProdMappingListNew',
'isValid' => '2',
'mappingState' => '0',
'noMap' => '0',
'reportName' => 'server/dmsProdMappingListNew',
'server_dmsProdMappingListNew_crd' => '99999',
'server_dmsProdMappingListNew_p' => '1',
'server_dmsProdMappingListNew_rd' => '99999',
'totalstatus' => '2',
};

my  $tx=$ua->post($post_url => form => $post_form );
if (encode("gb2312",H($tx->res->body)) =~/区/){
    my @items=encode("gb2312",H($tx->res->body)) =~ m/区/g;
    say " -> 中区有 ",scalar @items," 条记录";

	# 直接从网页中生成 xlsx 文件
	 my $workbook   =  Excel::Writer::XLSX->new( "产品匹配维护.xlsx" );
     my $new_worksheet  =  $workbook->add_worksheet('PV');
	 my @columns;
    
     my $collection = $tx->res->dom->find('tbody > tr');  # 取得所有列，每个 tr 是一列
     # say $collection;
     foreach my $ele (@$collection) { 
       my $td = $ele->find('td')->text; # 每一行中的所有元素
	   $td =~s/&#160;//g;
	   # say encode("gb2312",$td);
       push @columns,[@$td] if encode("gb2312",$td) =~/中区/;
	   # <>;
     }
	 
my $count= scalar @columns;
$new_worksheet->keep_leading_zeros(); # 保留前导0	 
$new_worksheet->write_col( 'A1', \@columns);
say " -> download $count rows";	
exit;	
} else {
  say " -> [ OK NOW ]";
  say " ---------------";
  }
}
}
}
}
);
$cv->recv;