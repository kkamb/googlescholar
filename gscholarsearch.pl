#!/usr/bin/perl -w

###################################################
#GETTING INFORMATION FROM GOOGLE SCHOLAR
###################################################
#needs input file GoogleScholar.xls with url in fourth column
#outputs new GoogleScholar2.xls file with:
#0/1 whether User Profile or not in 1st column;
#user profile link in 2nd column if User Profile exists (otherwise 0)
#total citation no in 3rd column if User Profile exists (otherwise 0)
#####################################################

use Spreadsheet::ParseExcel;
use Spreadsheet::WriteExcel;
use LWP::UserAgent;

$Excelin="GoogleScholar.xls";
$Excelout="GoogleScholar2.xls";
$urlcolno=3;

####reading in url data from excel spreadsheet
####creating a @googlescholarurl array to store all the urls

my $parser = Spreadsheet::ParseExcel->new();
my $wbookin = $parser->Parse($Excelin);
my @googlescholarurl = ();
my $workbook = Spreadsheet::WriteExcel->new($Excelout);
my $linksheet = $workbook->add_worksheet("htmllinks");

for my $worksheet ( $wbookin->worksheets() ) {
    my ( $row_min, $row_max ) = $worksheet->row_range();
    for my $row ( 1 .. $row_max ) {
        my $cell = $worksheet->get_cell($row, $urlcolno );
        next unless $cell;
        my $trialno = $cell->value();
        push(@googlescholarurl,$trialno);
    }
}


####going through all the urls one by one, grabbing citation and pagename data if User Profile exists on that page
####storing citations data in @gpagecit, whether User Profile exists or not in @gpageyesno
####and google unique url in @gpagename
###urls of the form http://scholar.google.com/scholar?hl=en&q=Linda+DeAngelo+Southern+California
#https://scholar.google.com/citations?view_op=search_authors&mauthors=malcolm+baker+harvard&hl=en&oi=ao

my $authorno=@googlescholarurl; #length of googlescholarurl array
@gpagecit = @gpagename = @gpageyesno = (0) x $authorno;

for($currentno=0; $currentno<$authorno; $currentno++){

    # Create a user agent object
    my $ua = LWP::UserAgent->new(env_proxy=>1,agent=>"Mozilla/34.0.5 ");

    # Create a request
    my $base=$googlescholarurl[$currentno];
    my $res  = $ua->request( HTTP::Request->new( GET => $base ) );
 
    # Check the outcome of the response
    if ($res->is_success) {
        my $content=$res->content;
        my $outputfile='tempresult4.txt';
        open (OUTFILE, ">$outputfile") or die "Could not open file '$outputfile' $!";
            print OUTFILE $content;
        close (OUTFILE);
    
        open(MYINFILE, $outputfile) or die "Could not open file '$inputfile' $!";
            @cont = <MYINFILE>;
        close(MYINFILE);

        # and delete it
        unlink($outputfile);
        my $text;

        # append all the lines, in order to remove the carriage returns and new lines
        foreach my $line (@cont){
            chomp($line);
            $text=$text.$line;
        }
 
        # look through the html code and find the user id and citation number
        if($text =~ m/(User profiles for)/){
            $gpageyesno[$currentno]=1;
            if($text =~ m/(citations\?user=............&amp;hl=en&amp;oi)/ ){
                $gpagename[$currentno]="http://scholar.google.com/" . $1;
            }
            if($text =~ m/(<div>Cited by \w+)/ ){
                $gpagecit[$currentno]=$1;    
                substr($gpagecit[$currentno],0,13)=""; #get rid of the first 13 characters (<div>Cited by )
            }
        }
        
        #store output in excel file
        $linksheet->write($currentno+1,1,$gpagename[$currentno]);
        $linksheet->write($currentno+1,0,$gpageyesno[$currentno]);
        $linksheet->write($currentno+1,2,$gpagecit[$currentno]);
    }
    else{
        print "Error: Could not get to Google Search Page \n";
    }
}