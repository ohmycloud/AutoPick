#!/usr/bin/perl
use strict;
use warnings;

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

# or using the eval{ $obj->Method()->Method()->...->Close()} trick ...
        use Mail::Sender;
        eval {
        (new Mail::Sender)
                ->OpenMultipart({
                            smtp => 'smtp.winchannel.net',
                            from => 'sunxiaowei@winchannel.net',
                             auth=> 'LOGIN',
                          authid => 'sunxiaowei@winchannel.net',
                         authpwd => 'win123456',
                              to => 'sxw2k@sina.com',
                         subject => "$today工作日结",
                        boundary => 'boundary-test-1',
                            type => 'multipart/related'
                })
                ->Attach({
                        description => 'fujian',
                              ctype => 'application/x-zip-encoded',
                           encoding => 'Base64',
                        disposition => 'NONE',
                               file => "E:/GSK-RX远程登记表.xlsx,E:/上传跟进情况$today.xlsx,E:/工作日志.xlsx"
                })
           
                ->Close()
        }
        or die "Cannot send mail: $Mail::Sender::Error\n";