use 5.010;
use strict;  
# use warnings;  
use Spreadsheet::XLSX;  
use Excel::Writer::XLSX;
use MyExcelFormatter;
use Encode;
use HTML::TokeParser;
use Data::Dumper;
use Mojo::UserAgent;
use Mojo::UserAgent::CookieJar;
use Mojo::UserAgent::Proxy;
use YAML 'Dump';
use Win32::API;

#获取当天的日期，作为后面 Excel 的行首
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

my $date = getTime();  
my $today = $date->{date};                          # 获取xxxx-xx-xx这样的日期
#my $month = $date->{month};                        # 获取月
#my $day = $date->{day};                            # 获取日
#my $year = $date->{year};                          # 获取年
#my $weekday = $date->{wday};                       # 获取星期
my $yesterday = getTime(time() - 86400)->{date};   # 获取昨天的日期，也可以用 86400*N，获取N天前的日期

sub H{
my $text = shift;
return  decode('utf8',$text);  # 进行转码
}
##################################################
#                 删除旧文件                     #
##################################################

foreach my $file (grep {/客户端运行日志/ or /运维操作平台/} glob "*.xls") { 
say "删除文件：$file";
unlink  "$file";
}

my $client_log_file = "客户端运行日志".$today.".htm.xls";
my $yesterday_file  = 'Finished上传跟进情况'.$yesterday.'.xlsx';
my $yunwei_file     = "运维操作平台".$today.".htm.xls";

##################################################
#                 下载日志文件                    #
##################################################

my $proxy = Mojo::UserAgent::Proxy->new;
$proxy->detect;
say $proxy->http;


# 需要下载 AdvOcr

# 这个版本是可以的，因为有时候要输入验证码，所以有时不行
my $ua   = Mojo::UserAgent->new;
my $http = $proxy->http;
my $ua_proxy      = $proxy->http('http://192.168.1.158:8080');
my $response = $ua->get('https://gskrx.windms.com/commons/image.jsp');

 if ($response->success) {
     # 抓取验证码图片 #
$response->res->content->asset->move_to('verifycode.BMP'); # 不能含中文？


 # # 输入验证码并登录 #
# print "--> enter verifycode:";

# chomp( my $verifycode = <> );
# $verifycode=~ s/\s//g; #删除空白

my $dll={};
my $D='AdvOcr.dll';
$dll->{OcrInit} = Win32::API->new($D, 'OcrInit',[],'N') || die " Can't open the dll file $D $!";

$dll->{OCR_C} = Win32::API->new($D, 'OCR_C',['P','P'],'P');
if($dll->{OcrInit}->Call()){
my $ocr_txt=$dll->{OCR_C}->Call('163_esales','verifycode.BMP');
    print "结果:$ocr_txt\n";
# }

# 163_esales 就是使用哪个字模库
# 作者做了很多现成的字模，只要更换这个参数，就可以识别其他类型的验证码
# 具体参数，看OCRtypedef.ini文件，里面有28种字模

my $login_url  = "https://gskrx.windms.com/login.do";
my $post_form = {
       'rand' => "$ocr_txt",
       'anchor'       =>"",
       'userAccount'  =>'winc_sxw',
       'userPassword' =>'000000',

    };

  
    my $tx = $ua->post( $login_url => form => $post_form );
if ( $tx->success ) {
    if ($tx->res->body !~ /验证码输入错误/) {
	
my $query_form= {
'colIds' => '0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,48,43,44,45,46,47',
'reportName' => '../listview/platform/server/dmsRunLogListView',
'tableId'=> 'dmsRunLogListView',
'uploadResult'=>'2',
'loadResult'=>'2',
'uploadDate1'=>"$today",
'uploadDate2'=>"$today",
'dmsRunLogListView_rd'=>'200',
'_RES_ID_'=>'156',
'ec_i'=>'dmsRunLogListView',
'ec_eti'=>'dmsRunLogListView',
'dmsRunLogListView_ev'=>'htm',
'dmsRunLogListView_efn'=>'客户端运行日志.htm.xls',
'dmsRunLogListView_crd'=>'200',
'dmsRunLogListView_p'=>'1',
};

my $yunwei_query_form = {
'_RES_ID_'=>'236',
'colIds'=>'0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,36,37',
'delayDay'=>'2',
'ec_eti'=>'helpDesk',
'ec_i'=>'helpDesk',
'FILE_EXPORT_TOKEN_PARAM_NAME'=>'1393911174875',
'helpDesk_crd'=>'200',
'helpDesk_efn'=>'运维操作平台.htm.xls',
'helpDesk_ev'=>'htm',
'helpDesk_p'=>'1',
'helpDesk_rd'=>'200',
'reportName'=>'helpDesk',
'saleDelayDay'=>'0',
};

# 'FILE_EXPORT_TOKEN_PARAM_NAME'=>'1392799287157'
# URL 要经过解码才能POST
my $client_file='客户端运行日志'.$today.'.htm.xls';

my  $tx=$ua->post('https://gskrx.windms.com/report/listReport.do' => form => $query_form );
$tx->res->content->asset->move_to($client_file); # 不能含中文？

my $yunwei_tx =$ua->post('https://gskrx.windms.com/report/listReport.do' => form => $yunwei_query_form);
$yunwei_tx->res->content->asset->move_to($yunwei_file); #

}
}
}
}
##################################################
#                检测是否下载到                  #
##################################################


my @files_needed =($client_log_file,$yesterday_file,$yunwei_file);
foreach my $file (@files_needed){
   print "[不存在] -> $file \n" if not -e $file;
}

exit   if not -e "运维操作平台".$today.".htm.xls";
exit   if not -e "客户端运行日志".$today.".htm.xls";
exit   if not -e 'Finished上传跟进情况'.$yesterday.'.xlsx';

my $parser = HTML::TokeParser->new($yunwei_file)
    or die "Can't open $yunwei_file: $!\n";
	
	
my (@table, @row, $inrow);
while (my $token = $parser->get_token( )) {
    my $type = $token->[0];
    if ( $type eq 'T' ) {
        push @row, $token->[1] if $inrow;
    }
    elsif ( $type eq 'S' ) {
        if ( $token->[1] eq 'tr' ) {
            $inrow = 1;
        }
    }
    elsif ( $type eq 'E' ) {
        if ( $token->[1] eq 'tr' ) {
            push @table, [@row]; # 注意这一行不能用 @row
            @row = ();
            $inrow = 0;
        }
    }
}
my @temp;
my %client_code;
foreach my $ele (@table) {
   my $string=join "", $ele->[10],$ele->[12],$ele->[14],"\n";
   $string=~s/\s+$//g;
   # print $ele->[3];
   push @temp, $string if $string;
   $client_code{$ele->[3]} =  $string;
}
delete $client_code{'客户端编码'};
# print Dumper(%client_code);



################################################################################################
# 这里先提取客户端运行日志里的经销商编码，作为 Excel中 Vlookup函数的 table_array
################################################################################################

################################################################################################
# 这里需要准备的文件有3个，昨天的上传跟进情况，今天的客户端运行日志
################################################################################################


open my $fh,"<",$client_log_file or die $!;  #客户端运行日志2014-01-05.htm.xls
my @array;
my %has_seen;
while(<$fh>) {
     $has_seen{$1}=1 if /<td class="tsc" >(\d{7,})<\/td>/;
}

foreach my $key (keys %has_seen) {
        push @array,$key;
}
close $fh;

my $workbook   =  Excel::Writer::XLSX->new( "上传跟进情况".$today.".xlsx" );
my $format     = $workbook->add_format( 
           align      => 'center',
		   font       => H('微软雅黑'),
		   size       => 9 ,
		   num_format => '@' 
		   );
		   
my $time_format     = $workbook->add_format( align => 'center', num_format => 'h:mm',        font => H('微软雅黑'), size => 9);
my $date_format     = $workbook->add_format( align => 'center', num_format => 'yyyy/mm/dd',  font => H('微软雅黑'), size => 9);
my $filter_format_Y = $workbook->add_format( align => 'center', bg_color   => '#16a951');
my $filter_format_N = $workbook->add_format( align => 'center', bg_color   => 'red', );

# set_column( $first_col, $last_col, $width, $format, $hidden, $level, $collapsed )
   
# 读取一个Excel 的内容到另外一个新建的工作簿中，其实就是复制原来的 Excel 文件	   
my $excel          =  Spreadsheet::XLSX -> new ('Finished上传跟进情况'.$yesterday.'.xlsx');

foreach my $old_sheet( @{$excel -> {Worksheet}}[0] ) {  # 只读取前两张工作表，直链安装
        my @columns; 
        my $new_worksheet  =  $workbook->add_worksheet(H($old_sheet->{Name}));  # 将原来的工作表名添加到新的 Excel
	    $old_sheet -> {MaxCol} ||= $old_sheet -> {MinCol};  #|| 逻辑或，返回计算结果先为真的值，由左到右计算
        my $temp=$old_sheet -> {MaxCol}+65;  # 今天最大列数是10，今天结束会增加一列，列数变成11,自增1,将数字转换为对应的字母
        my $cell_today=chr($temp+1);

	   # 写入新 Excel 文件前，先设置新 Excel 的单元格格式 
       $new_worksheet->set_column( 'A:A', 8.38,$format );
       $new_worksheet->set_column( 'B:B', 10.38,$format);
       $new_worksheet->set_column( 'C:D', 25 ,$format);
       $new_worksheet->set_column( 'F:F', 15, $time_format );
	   $new_worksheet->set_column( 'G:G', 25 ,$date_format);
       $new_worksheet->set_column( 'E:E', 15, $format ); 
       $new_worksheet->set_column( "J:$cell_today", 15, $date_format ); 
       
# 读取原 Excel，然后写入 新 Excel        
foreach my $col ($old_sheet -> {MinCol} .. $old_sheet -> {MaxCol}) {
			    $old_sheet -> {MaxRow} ||= $old_sheet -> {MinRow}; 
                foreach my $row ($old_sheet -> {MinRow} ..  $old_sheet -> {MaxRow}) {  
					    push @{$columns[$col]},H($old_sheet -> {Cells} [$row] [$col]->{Val});
                        }
				 $new_worksheet->write_row( 'A1', \@columns);
				}


$new_worksheet->write( $cell_today."1",$today,$date_format); # header，为每天的日期

shift  @{$columns[1]}; 
chomp @{$columns[1]};
my @custom=map {s/\s+//g;$_} @{$columns[1]};
shift @{$columns[4]};
my @source_client =map {s/\s+//g;$_} @{$columns[4]};
{no warnings;
foreach my $i ($old_sheet -> {MinRow}+2..$old_sheet -> {MaxRow}+1) {
        $new_worksheet->write( 'F'.$i, $source_client[$i-2] ~~ %client_code ? decode("gb2312",$client_code{$source_client[$i-2]}):"N");
	}

foreach my $i ($old_sheet -> {MinRow}+2..$old_sheet -> {MaxRow}+1) {
        $new_worksheet->write( $cell_today.$i, $custom[$i-2] ~~ @array ? "Y":"N");
	}

my $alpha=$cell_today;
my $b=$alpha.($old_sheet -> {MaxRow}+1);	
$new_worksheet->autofilter( "A1:$b" ); # 对选区数据进行筛选
$new_worksheet->filter_column( 8, 'x == NonBlanks');
# $new_worksheet->filter_column( 5, 'x =~ *小时');
    		
my $row = 1;
shift @{$columns[8]};
    
for my $region ( @{$columns[8]} ) {
        if ( not defined $region) {
    
            # Hide row.
            $new_worksheet->set_row( $row, undef, undef, 1 ); # 隐藏不匹配过滤标准的行
        } 
    
        $new_worksheet->write( $row++, 8, $region );
		$new_worksheet->write( $cell_today.$row,$region ~~ @array ? "Y":"N") if defined $region;# 最后的if defined 一定要加上，这是关键
 }
}  
		 		
    #条件格式 
    		  $new_worksheet->conditional_formatting( "J:$cell_today",
                {
                    type     => 'text',
                    criteria => 'begins with',
                    value    => 'Y',
                    format   => $filter_format_Y,
                }
            );
    		
    		  $new_worksheet->conditional_formatting( "J:$cell_today",
                {
                    type     => 'text',
                    criteria => 'begins with',
                    value    => 'N',
                    format   => $filter_format_N,
                }
            );
}
print  "---------> [ OK ]";
__END__