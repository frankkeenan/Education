#!/usr/bin/perl
use Getopt::Std;
use autodie qw(:all);
use utf8;
use Excel::Writer::XLSX;
use strict;
our ($LOG, $LOAD, $opt_r, $opt_f, $opt_u, $opt_D, $opt_I, $opt_O);
our ($MAXROW, %TOPLEVS);
#
require "/usrdata3/dicts/NEWSTRUCTS/progs/utils.pl";
require "/NEWdata/dicts/generic/progs/restructure.pl";
#require "/NEWdata/dicts/generic/progs/xsl_lib_fk.pl";
$LOG = 1;
$LOAD = 0;
$, = ' ';               # set output field separator
$\ = "\n";              # set output record separator
#undef $/; # read in the whole file at once
our ($workbook, $worksheet);
our ($unlocked, $locked, $hidden, $format1, $fmt_tint, $fmt_wrap);
our($opt_u);

&main;

sub main
{
    getopts('uf:L:IODr:');
    &usage if ($opt_u);
    my($e, $res, $bit);
    my(@BITS);
    #   $opt_L = ""; # name of file for the log_fp output to go to
    use open qw(:std :utf8);
    unless ($opt_r)
    {
	$opt_r = "fk_test.xlsx";
    }
    &open_debug_files;
    if ($opt_D)
    {
	binmode DB::OUT,":utf8";
    }
    $workbook  = Excel::Writer::XLSX->new( $opt_r );
    $worksheet = $workbook->add_worksheet();
    &do_formats;
    my $row = 0;
    my @HDR = ("Headword", "homograph", "form", "pos", "sense number", "definition", "examples");
    my $HDR_ref = \@HDR;
    $worksheet->write_row($row++, 0, $HDR_ref);
  line:    
    while (<>){
	chomp;       # strip record separator
	s|||g;
	if ($opt_I){printf(bugin_fp "%s\n", $_);}
	##
	my $hw = restructure::get_tag_contents($_, "hw");
	my $hm = restructure::get_tag_attval($_, "hw", "homograph"); 
	$_ =~ s|(<msDict[ >].*?</msDict>)|&split;&fk;$1&split;|gi;
	my @BITS = split(/&split;/, $_);
	my $res = "";
	foreach my $bit (@BITS){
	    if ($bit =~ s|&fk;||gi){
		my $form = restructure::get_tag_attval($bit, "msDict", "form");
		my $pos = restructure::get_tag_attval($bit, "msDict", "pos");
		my $sensenum = restructure::get_tag_attval($bit, "msDict", "sensenum"); 
		my $def = restructure::get_tag_contents($bit, "df"); 
		my $exas = &get_examples($bit);
		$def =~ s|(<[^ >]*) [^>]*>|\1>|gi;
		my @E = ($hw, $hm, $form, $pos, $sensenum, $def, $exas);
		$worksheet->write_row($row++, 0, \@E);
	    }
	}
	if ($opt_O){printf(bugout_fp "%s\n", $_);}
  }
    $workbook->close();
    &close_debug_files;
}

sub get_examples
{
    my($e) = @_;
    my($res, $eid);	
    $e =~ s|(<ex[ >].*?</ex>)|&split;&fk;$1&split;|gi;
    my @BITS = split(/&split;/, $e);
    my $res = "";
    foreach my $bit (@BITS){
	if ($bit =~ s|&fk;||gi){
	    $bit = restructure::get_tag_contents($bit, "ex"); 
	    $res .= sprintf("%s &xsep; ", $bit); 
	}
    }
    $res =~ s| *&xsep; *$||;
    $res =~ s|&xsep;|◆|g;
    return $res;
}


sub do_senses
{
    my($e, $row) = @_;
    my($res, $eid);	
    $e =~ s|(<semb[ >].*?</semb>)|&split;&fk;$1&split;|gi;
    my @BITS = split(/&split;/, $e);
    my $res = "";
    foreach my $bit (@BITS){
	if ($bit =~ s|&fk;||gi){
	    if ($bit =~ m|<label|){
		my $levs = &get_levs($bit);
		my $labels = &get_labels($bit);
#		$worksheet->write( $row, 1, $levs, $fmt_wrap );
		$worksheet->write( $row, 2, $labels, $fmt_wrap );		
		printf(log_fp "r$row c1 [LEVS] $levs\n") if ($LOG); 
	    }


	    $bit =~ s| *<label[^>]*>(.*?)</label> *| \1 |gi;
	    $bit =~ s| *<reg[^>]*>(.*?)</reg> *| \1 |gi;
	    my $trans = restructure::get_tag_contents($bit, "trans-gs");
	    my $exas = restructure::get_tag_contents($bit, "x-gs");
	    my $xrefs = restructure::get_tag_contents($bit, "xr-gs");
	    $bit = restructure::tag_delete($bit, "x-gs");
	    $bit = restructure::tag_delete($bit, "xr-gs");
	    $bit = restructure::tag_delete($bit, "trans-gs");	    
	    my $def = restructure::get_tag_contents($bit, "semb");
	    my $tint = 0;
	    if ($def =~ s| *<lev.*?>(.*?)</lev> *| \1 |gi)
	    {
		$tint = 1;
	    }
	    $def =~ s|^ *||;
	    if ($tint)
	    {
		$worksheet->write($row, 3, $def, $fmt_tint);
	    } else {
		$worksheet->write($row, 3, $def, $fmt_wrap);
	    }
	    printf(log_fp "r$row c3 [DEF] $def\n") if ($LOG);
	    &do_trans($trans, $row);
	    &do_examples($exas, $row);
	    $row = $MAXROW + 1;
	}
    }    
}


sub do_trans
{
    my($e, $row) = @_;
    my($res, $eid);	
    $e =~ s|(<trg[ >].*?</trg>)|&split;&fk;$1&split;|gi;
    my @BITS = split(/&split;/, $e);
    my $res = "";
    foreach my $bit (@BITS){
	if ($bit =~ s|&fk;||gi){
	    $bit = restructure::tag_delete($bit, "tgr"); 
	    my $ts = restructure::get_tag_contents($bit, "trg");
	    my $tint = 0;
	    if ($ts =~ s| *<lev.*?>(.*?)</lev> *| \1 |gi)
	    {
		$tint = 1;
	    }
	    $ts =~ s|<.*?>| |gi;
	    $ts =~ s| +| |g;
	    $ts =~ s|^ +||g;
	    $ts =~ s| +$||g;
	    if ($tint)
	    {
		$worksheet->write($row, 5, $ts, $fmt_tint);
	    } else {
		$worksheet->write($row, 5, $ts, $fmt_wrap);
	    }
	    printf(log_fp "r$row c5 [TS] $ts\n") if ($LOG);
	    $row++;
	}
    }    
    $row--; # blank row added for next trans
    if ($row > $MAXROW)
    {
	$MAXROW = $row;
    } 
}


sub do_examples
{
    my($e, $row) = @_;
    my($res, $eid);	
    $e =~ s|(</tx-g>)(<tx-g )|\1</exg><exg >\2|gi; # Cheat to force new lines for multiple trans
    $e =~ s|(<exg[ >].*?</exg>)|&split;&fk;$1&split;|gi;
    my @BITS = split(/&split;/, $e);
    my $res = "";
    foreach my $bit (@BITS){
	if ($bit =~ s|&fk;||gi){
#	    $bit =~ s|</tx-g><tx-g[^>]*>|; |gi;
		while ($bit =~ s|</tx[^>]*><tx[^>]*>|; |gi){}
	    $bit =~ s| *<label[^>]*>(.*?)</label> *| \1 |gi;
	    my $tint = 0;
	    if ($bit =~ s| *<lev.*?>(.*?)</lev> *| \1 |gi)
	    {
		$tint = 1;
	    }
	    my $ex = restructure::get_tag_contents($bit, "ex");
	    my $tx;
	    if ($bit =~ m|<tx-g|)
	    {
		$tx = restructure::get_tag_contents($bit, "tx-g");
	    } else {
		$tx = restructure::get_tag_contents($bit, "tx");
	    }
	    $tx =~ s|<.*?>| |gi;
	    s| +| |g;
	    $worksheet->write($row, 7, $ex, $fmt_wrap);
	    $worksheet->write($row, 9, $tx, $fmt_wrap);
	    printf(log_fp "r$row c7 [EX] $ex\n") if ($LOG);
	    printf(log_fp "r$row c9 [TX] $tx\n") if ($LOG);
	    $row++;
	}
    }    
    $row--; # blank row added for next example
    if ($row > $MAXROW)
    {
	$MAXROW = $row;
    } 
}



sub remove_grambs_without_lev
{
    my($e) = @_;
    my($res, $eid);	
    $e =~ s|(<gramb[ >].*?</gramb>)|&split;&fk;$1&split;|gi;
    my @BITS = split(/&split;/, $e);
    my $res = "";
    undef %TOPLEVS;
    foreach my $bit (@BITS){
	if ($bit =~ s|&fk;||gi){
	    if ($bit =~ m|<lev|)
	    {
		&store_toplevs($bit);
	    } else {
		$bit = "";
	    }
	}
	$res .= $bit;
    }
    my $levs = "";
    foreach my $lev (sort keys %TOPLEVS)
    {
	$levs = sprintf("%s%s", $levs, $lev); 
    }
    $res =~ s|(</hwg>)|$levs$1|;
    return $res;
}

sub store_toplevs
{
    my($e) = @_;
    my($res, $eid);	
    $e =~ s|<semb.*||;
    $e =~ s|(<lev[ >].*?</lev>)|&split;&fk;$1&split;|gi;
    my @BITS = split(/&split;/, $e);
    my $res = "";
    foreach my $bit (@BITS){
	if ($bit =~ s|&fk;||gi){
	    $TOPLEVS{$bit} = 1;
	}
    }    
}



sub do_grambs
{
    my($e, $row) = @_;
    my($res, $eid);	
    $e =~ s|(<gramb[ >].*?</gramb>)|&split;&fk;$1&split;|gi;
    my @BITS = split(/&split;/, $e);
    my $res = "";
    foreach my $bit (@BITS){
	if ($bit =~ s|&fk;||gi){
	    my $pos = restructure::get_tag_contents($bit, "ps");
	    $worksheet->write($row, 4, $pos, $fmt_wrap);
	    printf(log_fp "r$row c4 [POS] $pos\n") if ($LOG);
	    &do_senses($bit, $row);
	    $row = $MAXROW+1;
	}
    }    
}

sub usage
{
    printf(STDERR "USAGE: $0 -u \n"); 
    printf(STDERR "\t-u:\tDisplay usage\n"); 
    #    printf(STDERR "\t-x:\t\n"); 
    exit;
}

sub get_labels
{
    my($e) = @_;
    my($res, $eid);	
    my %USED;
    $e =~ s|(<label[ >].*?</label>)|&split;&fk;$1&split;|gi;
    my @BITS = split(/&split;/, $e);
    my $res = "";
    foreach my $bit (@BITS){
	if ($bit =~ s|&fk;||gi){
	    my $label = restructure::get_tag_contents($bit, "label"); 
	    unless ($USED{$label}++)
	    {
		$res = sprintf("%s%s, ", $res, $label); 
	    }
	}
    }    
    $res =~ s|, *$||;
    return $res;
}

sub get_levs
{
    my($e) = @_;
    my($res, $eid);	
    my %USED;
    $e =~ s|(<lev[ >].*?</lev>)|&split;&fk;$1&split;|gi;
    my @BITS = split(/&split;/, $e);
    my $res = "";
    foreach my $bit (@BITS){
	if ($bit =~ s|&fk;||gi){
	    my $lev = restructure::get_tag_contents($bit, "lev"); 
	    unless ($USED{$lev}++)
	    {
		$res = sprintf("%s%s, ", $res, $lev); 
	    }
	}
    }    
    $res =~ s|, *$||;
    return $res;
}

sub get_lev_tags
{
    my($e) = @_;
    my($res, $eid);	
    my %USED;
    $e =~ s|(<lev[ >].*?</lev>)|&split;&fk;$1&split;|gi;
    my @BITS = split(/&split;/, $e);
    my $res = "";
    foreach my $bit (@BITS){
	if ($bit =~ s|&fk;||gi){
	    my $lev = restructure::get_tag_contents($bit, "lev"); 
	    unless ($USED{$lev}++)
	    {
		$res = sprintf("%s%s", $res, $bit); 
	    }
	}
    }    
    return $res;
}


sub do_formats
{    # Create some format objects
    $unlocked = $workbook->add_format( locked => 0 );
    $locked = $workbook->add_format( locked => 1 );
    $hidden   = $workbook->add_format( hidden => 1 );
    # Light red fill with dark red text.
    $format1 = $workbook->add_format(
	bg_color => '#E6FFFF',
	color    => '#9C0006',
	
	);
    
    # Green fill with dark green text.
    $fmt_tint = $workbook->add_format(
	bg_color => '#ECE75F',
	color    => '#006100',
	
	);
    $fmt_wrap = $workbook->add_format();
    $fmt_wrap->set_text_wrap();
    # Format the columns
    $worksheet->autofilter( 'A1:I9999' );
    $worksheet->freeze_panes( 1 );    # Freeze the first row
    $worksheet->set_column( 'A:A', 20, $unlocked );
    $worksheet->set_column( 'B:B', 20, $unlocked );
    $worksheet->set_column( 'C:C', 20, $unlocked );
    $worksheet->set_column( 'D:D', 20, $unlocked );
    $worksheet->set_column( 'E:E', 40, $unlocked );
    $worksheet->set_column( 'F:F', 80, $unlocked );
    $worksheet->set_column( 'G:G', 20, $unlocked );
    $worksheet->set_column( 'H:H', 80, $unlocked );
    $worksheet->set_column( 'I:I', 20, $unlocked );
    $worksheet->set_column( 'J:J', 80, $unlocked );
    $worksheet->set_column( 'K:K', 20, $unlocked );
    $worksheet->set_column( 'L:L', 80, $unlocked );
    $worksheet->set_column( 'M:M', 80, $unlocked );
    
    #    $worksheet->autofilter( 'A1:K1' );
    #    # Protect the worksheet
    #    $worksheet->protect("", {autofilter => 1});
    #    $worksheet->protect({autofilter => 1});
    #    protectWorksheet(wb, sheet = i, protect = TRUE, password = "Password") #Protect each sheet
}

