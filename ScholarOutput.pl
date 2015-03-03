#!/usr/bin/perl -w
use Spreadsheet::WriteExcel;
use HTML::Parser; 
use LWP::Simple; 
my $workbook = Spreadsheet::WriteExcel->new("GoogleScholar.xls");
my $linksheet = $workbook->add_worksheet("htmllinks");

my $inyear=1980;
my $outyear=2014;
my @scholarfname = my @scholarschool = my @scholarlname = ();
my @scholarstring = my @scholaryear = my @googlescholarurl = ();

for ($j=$inyear; $j<$outyear; $j++) {
    my $getstring = "http://fisher.osu.edu/fin/findir/indexAYG.html?gradYear=" . $j;
    my $html = get $getstring;
    HTML::Parser->new(text_h => [\my @accum, "text"])->parse($html);
    @arrayoflines= map("$_->[0] $_->[1]\n", @accum);
    ###outputs scholar name and university at @scholarstring
    my $count1=14;
    my @inputstring = ();
    until ($arrayoflines[$count1] !~ /(\w+)/g) {
        if ($count1 % 2) {
            @inputstring = split(' ', $arrayoflines[$count1]);
            $count2=1;
            if ($arrayoflines[$count1] =~ m/(at .*,)/ ){
                $schoolname=$1;
                substr($schoolname,0,3)="";
                substr ($schoolname, -1) = "";
            }
            push(@scholarschool,$schoolname);
            if ($inputstring[$count2] =~ m/University/) {
                $count2=$count2+2;
            }
            if ($inputstring[$count2] =~ m/(St.)|(San)/) {
                $schoolstring = $inputstring[$count2] . "+" . $inputstring[$count2+1];
            }
            else{
                $schoolstring = $inputstring[$count2];
            }
            $schoolstring =~ s/-/+/g;
            $schoolstring =~ s/[,]//g;
            $addstring = $namestring . "+" . $schoolstring;
            push (@scholarstring, $addstring);
            $urlcall = "http://scholar.google.com/scholar?hl=en&q=" . $addstring;
            push(@googlescholarurl,$urlcall);
            push(@scholaryear, $j);
        } else {
            $namestring = $arrayoflines[$count1];
            $namestring =~ s/ /+/;
            chomp $namestring;
            $namestring =~ s/ //;
            @nameinput = split('\+',$namestring);
            push (@scholarfname, $nameinput[0]);
            push (@scholarlname, $nameinput[1]);
        }
        $count1++;
    }
}

#output google search link (name of scholar and year) in excel
$csj=0;
foreach $scholar (@googlescholarurl){
    $linksheet->write($csj,0,$scholarfname[$csj]);
    $linksheet->write($csj,1,$scholarlname[$csj]);
    $linksheet->write($csj,2,$scholarschool[$csj]);
    $linksheet->write($csj,3,$googlescholarurl[$csj]);
    $linksheet->write($csj,4,$scholaryear[$csj]);
    $csj++;
}