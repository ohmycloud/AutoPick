#! /usr/bin/perl
# 学习perl LWP时用post做的翻译小脚本
# 调用的是有道词典
# 作者：第2012位菜鸟
# 原文链接 http://www.cnblogs.com/caibird/archive/2013/03/22/2974999.html
# 2014-01-10 增加编码判断+按CP936规范输出，适合在windows命令行显示 - paktc

use strict;
use warnings;
use LWP::UserAgent;
use JSON;
use Encode;
use Term::ANSIColor;
use Win32::Console::ANSI;
use 5.010;
my $mode=recogn();
binmode(STDOUT,':encoding(CP936)');

my $browser = LWP::UserAgent->new();
while(1){
print "> Please input the word:";
chomp (my $input = <STDIN>);
my $response = $browser->post(
#    'http://fanyi.youdao.com/translate?smartresult=dict&smartresult=rule&smartresult=ugc&sessionFrom=https://www.google.com.hk/',
    'http://fanyi.youdao.com/translate?smartresult=dict&smartresult=rule&smartresult=ugc&sessionFrom=null',
    [
        'type'    => 'AUTO',
        'i'       => "$input",
        'doctype' => 'json',
    ],
    );

if($response->is_success){
    my $result =eval{ $response->content };
    my $json = new JSON;
	my $obj;
    eval{
	$obj = $json->decode($result);
	};
    #print Dumper $obj;
    my $trans =eval{ @{$obj->{'translateResult'}[0]}[0]->{"tgt"} };
    my $string;
    eval{
        $string  = join " ", @{$obj->{'smartResult'}->{"entries"}};
    };

    my $say1=decode($mode," -> 翻译结果：");
    my $say2=decode($mode," -> 其他结果: ");
    $trans=decode('UTF-8',$trans) if $trans;
    $string=decode('UTF-8',$string) if $string;
    print color 'bold green';
    print $say1 if defined $trans;
    print color 'bold yellow';
	say $trans if defined $trans;
	print color 'bold red';
	print $say2 if defined $string;
	print color 'bold cyan';
	say $string if defined $string;
	print color 'bold white';
}
}
# 判断当前脚本编码格式  仅在WIN32中文系统测试过
sub recogn {
    my $cn="中";
    my $code;
    my @arr=split("",$cn);
    if ($#arr == 0) {            # 'Unicode' 4e2d
        $code='Unicode';
    } elsif ($#arr == 1) {       # 'GBK'   d6 d0 
        $code='GBK';
    } elsif ($#arr == 2) {       # 'UTF-8' e4 b8 ad
        $code='UTF-8';
    } else {
        $code='WHAT?';
    }
}