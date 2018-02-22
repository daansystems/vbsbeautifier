#!/usr/bin/perl
# vbscript beautifier
# (C)2001 Niek Albers, DaanSystems
# http://www.daansystems.com

use strict;
use Getopt::Std;
my $version  = "1.11";
my $location = '.';

if ( !@ARGV ) {
    print STDERR header();
    print STDERR qq(
Homepage: http://www.daansystems.com
Comments/Bugs: nieka\@daansystems.com
------------------------------------------
Usage: vbsbeaut [options] [files]

options:
 -i         Use standard input (as text filter).
 -s <val>   Uses spaces instead of tabs.
 -u         Make keywords uppercase.
 -l         Make keywords lowercase.
 -n         Don\'t change keywords.
 -d         Don\'t split Dim statements.

files:
 filenames  Wildcards allowed.
------------------------------------------
);
    exit;
}
my %options;
getopts( 'dr:iulns:', \%options );
print STDERR header();
if ( $options{'i'} ) {
    undef $/;
    my $text = <STDIN>;
    $/ = "\n";
    print do_all($text);
}
else {
    foreach my $filestring (@ARGV) {
        beautify($filestring) if ( -T $filestring );
    }
}

sub header {
    return qq(------------------------------------------
VBScript Beautifier v$version
(C)2001-2009 By Niek Albers - DaanSystems
------------------------------------------
);
}

sub beautify {
    my $filename = $_[0];
    my $text = read_file($filename) || die "ERROR: Can\'t open $filename.\n\n";
    my $beautified = do_all($text);
    write_file( "$filename.bak", $text );
    write_file( $filename, $beautified );
}

sub do_all {
    my $input = $_[0];
    my @inputlines = split( /\n/, $input );
    print STDERR "- Searching clientside VBScript delimiters.\n";
    replace_client_vbscript_tags( \@inputlines );
    $input = join( "\n", @inputlines );
    print STDERR "- Searching serverside VBScript delimiters.\n";
    my @html = gethtml( \$input );
    print STDERR "- Searching comments in clientside VBScript.\n";
    my @get_client_comments = getclientcomments( \$input );
    @inputlines = split( /\n/, $input );
    print STDERR "- Searching quoted text.\n";
    my @quoted = getquoted( \@inputlines );
    print STDERR "- Searching VBScript comments.\n";
    my @comments = getcomments( \@inputlines );

    #    print STDERR "- Splitting If..Then..Else on one line.\n";
    #    newline_after_then( \@inputlines );
    $input = join( "\n", @inputlines );

    #    preprocess( \$input );
    get_on_error_resume_next( \$input );
    @inputlines = split( /\n/, $input );
    if ( !$options{'d'} ) {
        print STDERR "- Splitting Dim statements.\n";
        preprocess2( \@inputlines );
    }
    $input = join( "\n", @inputlines );
    @inputlines = split( /\n/, $input );
    print STDERR "- Adjusting spaces around operators.\n";
    fixspaces( \@inputlines );
    if ( !$options{'n'} ) {
        print STDERR "- Modifying keywords.\n";
        replacekeywords( \@inputlines );
    }
    print STDERR "- Processing indent.\n";
    processindent( \@inputlines );
    print STDERR "- Lining out assignment statements.\n";
    lineoutequalsigns( \@inputlines );
    $input = join( "\n", @inputlines );
    put_on_error_resume_next( \$input );
    print STDERR "- Removing redundant newlines.\n";
    removeredundantenters( \$input );
    put_function_comments_to_declaration( \$input );
    putcomments( \$input, \@comments );
    putquoted( \$input, \@quoted );
    putclientcomments( \$input, \@get_client_comments );
    puthtml( \$input, \@html );
    $input =~ s/^\n\n+/\n/;
    replace_client_vbscript_tags_back( \$input );
    print STDERR "- All Done!\n";
    print STDERR "------------------------------------------\n";
    return $input;
}

sub newline_after_then {
    my $lines = $_[0];
    foreach my $line (@$lines) {
        $line =~ s/\s*then\s*/ then\n/gi;    # set enters behind then statements
    }
}

sub preprocess {
    my $text = $_[0];

#    $$text =~ s/\belse\b/\nelse\n/gi;
#    $$text =~ s/(?<!\bcase\b)\selse\b/\nelse\n/gi;
#    $$text =~ s/\belseif\b/\nelseif/gi;
#    $$text =~ s/end if/\nend if\n/gi;                # set enters behind end if statements
#    $$text =~ s/\n\s*\n*/\n/gs;                      # remove extra enters
}

# get html for server side vbscript (asp)
sub gethtml {
    my $text = $_[0];
    my ($starthtml) = $$text =~ m/^(.*?<%)/gs;    # find first html part
    $$text =~ s/^(.*?<%)/%[html]%/gs;    # substitute with template variable.
    my ($endhtml) = $$text =~ m/.*(%>.*)$/gs;    # find last html part
    $$text =~ s/(.*)(%>.*)$/$1%[html]%/gs;  # substitute with template variable.
    my @html = $$text =~
      m/(%>.*?<%)/gs; # return array of everythinh between >% and <% (aspswitch)
    $$text =~ s/(%>.*?<%)/%[html]%/gs;    # substitute with template variable.
    unshift( @html, $starthtml );
    push( @html, $endhtml );
    return @html;
}

# place back server side vbscript (asp)
sub puthtml {
    my $text    = $_[0];
    my $html    = $_[1];
    my $counter = 0;
    $$text =~ s/%\[html\]%/$$html[$counter++]/gse;
    $$text =~ s/(\S+)\s*%>/$1 %>/gs;
    $$text =~ s/<%\n/%[extraenter]%/gs;
    $$text =~ s/<%\s*(\S+)/<% $1/gs;
    $$text =~ s/%\[extraenter\]%/<%\n/gs;
}

sub replace_client_vbscript_tags {
    my $lines    = $_[0];
    my $count    = 0;
    my $found    = 0;
    my $endfound = 0;
    foreach my $line (@$lines) {
        if ( $found == 0 ) {
            $found =
              $line =~ s{(<\s*script.*?vbscript.*?>)}{$1 %[clientscript]%<%}i
              if ( $line !~ m{".*?<\s*script.*?vbscript.*?>.*?"}i );
        }
        else {
            $endfound = $line =~ s{(</\s*script\s*>)}{%>%[clientscript]% $1}i
              if ( $line !~ m{".*?</\s*script\s*>.*?"}i )
              ;    # substitute with template variable.
            $found = 0 if ( $endfound != 0 );
        }
    }
}

sub replace_client_vbscript_tags_back {
    my $text = $_[0];
    $$text =~ s/%>%\[clientscript\]%/\n/gs;
    $$text =~ s/%\[clientscript\]%<%/\n/gs;
}

sub getclientcomments {
    my $text     = $_[0];
    my @comments = $$text =~ m{(<!--|-->)}g;    # return all comments
    my $count    = $$text =~
      s{(<!--|-->)}{%[clientcomments]%}g;   # substitute with template variable.
    die "ERROR: Clientside comments not balanced!\n\n"
      if ( ( $count % 2 ) != 0 );
    return @comments;
}

sub putclientcomments {
    my $text    = $_[0];
    my $html    = $_[1];
    my $counter = 0;
    $$text =~ s/%\[clientcomments\]%/$$html[$counter++]/gse;
}

sub getcomments {
    my $lines = $_[0];
    my @allcomments;
    foreach my $line (@$lines) {
        my @comments =
          $line =~ m/(\'.*)/;    # return array of everything that is a comment
        $line =~ s/(\'.*)/%[comment]%/;    # substitute with template variable.
        push( @allcomments, @comments );
    }
    return @allcomments;
}

sub putcomments {
    my $text     = $_[0];
    my $comments = $_[1];
    my $counter  = 0;
    $$text =~ s/%\[comment\]%/$$comments[$counter++]/gse;
}

sub getquoted {
    my $lines = $_[0];
    my @allquoted;
    foreach my $line (@$lines) {
        my @quoted =
          $line =~ m/(".*?")/g;    # return array of everythinh between ""
        $line =~ s/(".*?")/%[quoted]%/g;    # substitute with template variable.
        push( @allquoted, @quoted );
    }
    return @allquoted;
}

sub putquoted {
    my $text    = $_[0];
    my $quoted  = $_[1];
    my $counter = 0;
    $$text =~ s/%\[quoted\]%/$$quoted[$counter++]/gse;
}

sub preprocess2 {
    my $lines = $_[0];
    foreach my $line (@$lines) {
        if ( $line =~ m/\bdim\b/i
          ) # replace all occurrances of comma separated dim variables on new lines
        {
            $line =~ s/,\s*/\ndim /g;
        }
    }
}

sub fixspaces {
    my $lines = $_[0];
    foreach my $line (@$lines) {
        $line =~ s/^\s*(.*?)\s*$/$1/;    # strip leading and trailing spaces
        $line =~ s/\s*(=|<|>|-|\+|&)\s*/ $1 /g;    # add spaces around signs
        $line =~ s/\s*<\s*>\s*/ <> /g; # remove spaces around and in  <> signs
        $line =~ s/\s*<\s*=\s*/ <= /g; # remove spaces around and in  <= signs
        $line =~ s/\s*=\s*<\s*/ =< /g; # remove spaces around and in  =< signs
        $line =~ s/\s*>\s*=\s*/ >= /g; # remove spaces around and in  >= signs
        $line =~ s/\s*=\s*>\s*/ => /g; # remove spaces around and in  => signs
        $line =~ s/\s*!\s*=\s*/ != /g; # remove spaces around and in  != signs
        $line =~ s/\s*<\s*%\s*/<% /g;  # remove spaces around <% signs
        $line =~ s/\s*%\s*>/ %>/g;     # remove spaces around %> signs
        $line =~ s/\s*_\s*$/ _/;       # add space before _ sign at end of line.
    }
}

sub countdelta {
    my $line     = $_[0];
    my $keywords = $_[1];
    my $indents  = $_[2];
    my $delta    = 0;
    foreach my $keyword (@$keywords) {
        $delta += $$indents{$keyword}
          if ( $$line =~ s/\b$keyword\b//gi );    # subtract closers
    }
    return $delta;
}

sub wordcount {
    my ($line) = @_;
    my $count = () = $line =~ /\w+/g;
    return $count;
}

sub get_keywords_indent {
    my ($section) = @_;
    open( FILE, "$location/keywords_indent.txt" )
      || die "Can\'t open $location/keywords_indent.txt";
    my @keywords            = ();
    my @singleline_keywords = ();
    my %indents;
    foreach my $line (<FILE>) {
        $line =~ s/\015?\012?$//;
        next if ( $line =~ m/^\s*$/ );
        next if ( $line =~ m/;/ );       # this is a line with comments....
        my ( $indent, $words, $singleline ) = split( /,/, $line );
        $indents{$words} = $indent;
        if ( !$singleline ) {
            push( @keywords, $words );
        }
        else {
            push( @singleline_keywords, $words );
        }
    }
    close(FILE);
    my @keywords_sorted =
      sort { wordcount($main::b) <=> wordcount($main::a) } @keywords;
    my @singleline_keywords_sorted =
      sort { wordcount($main::b) <=> wordcount($main::a) } @singleline_keywords;
    return ( \@keywords_sorted, \@singleline_keywords_sorted, \%indents );
}

sub processindent {
    my $lines  = $_[0];
    my $spaces = "\t";
    $spaces = ' ' x $options{'s'} if ( $options{'s'} );
    my $tabtotal = 0;
    my ( $indentors, $singleline_indentors, $indents ) = get_keywords_indent();
    foreach my $line (@$lines) {
        my $delta       = 0;
        my $singledelta = 0;
        my $linecopy    = $line;
        my $olddelta    = $delta;
        singlelineifthen( \$linecopy );
        $delta += countdelta( \$linecopy, $indentors, $indents );
        $singledelta -=
          countdelta( \$linecopy, $singleline_indentors, $indents );
        $tabtotal +=
          ( $delta < 0 ) ? $delta : 0;    # subtract closing braces/parentheses
        my $i = ( $tabtotal > 0 ) ? $tabtotal : 0;    # create tab index
        $tabtotal +=
          ( $delta > 0 )
          ? $delta
          : 0;    # add opening braces/parentheses for next print
        $line = $spaces x ( $i - $singledelta ) . $line;
        $line = "\n" . $line if ( $delta > $olddelta );
        $line = $line . "\n" if ( $delta < $olddelta );
    }

    #    die "ERROR: Indentation error!\n\n" if ( $tabtotal != 0 );
}

sub replacekeywords {
    my $lines = $_[0];
    undef $/;
    open( FILE, "$location/keywords.txt" )
      || die "Can\'t open $location/keywords.txt" . "\n\n";
    my $keywordsstring = <FILE>;
    close(FILE);
    $/ = "\n";
    my @keywords = split( /\n/, $keywordsstring );
    foreach my $line (@$lines) {
        foreach my $keyword (@keywords) {
            $keyword =~ s/\015?\012?$//;
            $keyword = uc($keyword) if ( $options{'u'} );
            $keyword = lc($keyword) if ( $options{'l'} );
            $line =~
              s/(\b)$keyword(\b)/$1$keyword$2/gi;    # substitute whole keywords
        }
    }
}

sub removeredundantenters {
    my $input = $_[0];
    $$input =~ s/\n\s*\n+/\n\n/gs;
}

sub get_linenumbers_with_statements {
    my $lines = $_[0];
    my @statement_linenumbers;
    my $counter = 0;
    foreach my $line (@$lines) {
        push( @statement_linenumbers, $counter ) if ( is_assignment($line) );
        $counter++;
    }
    return @statement_linenumbers;
}

sub lineoutequalsigns {
    my $lines                 = $_[0];
    my @statement_linenumbers = get_linenumbers_with_statements( \@$lines );
    my $lastlinenumber        = 0;
    my @linenumberlist        = ();
    foreach my $linenumber (@statement_linenumbers) {
        if (
            (
                   $lastlinenumber == $linenumber - 1
                || $lastlinenumber == $linenumber - 2
            )
            && $lastlinenumber != 0
          )
        {
            push( @linenumberlist, $lastlinenumber );
            push( @linenumberlist, $linenumber );
        }
        else {
            if ( scalar(@linenumberlist) >= 2 ) {
                lineout( $lines, \@linenumberlist );
            }
            @linenumberlist = ();
        }
        $lastlinenumber = $linenumber;
    }
}

sub is_assignment {
    my $line = $_[0];
    return 1 if ( $line =~ m/^\s*(set)?\s*\S+\s*=.*$/i );
    return 0;
}

sub lineout {
    my $lines          = $_[0];
    my $linenumbers    = $_[1];
    my $last_equal_pos = get_max_equal_pos_from_lines( $lines, $linenumbers );
    foreach my $linenumber (@$linenumbers) {
        my $equalsign = index( @$lines[$linenumber], '=' );
        my $diff = $equalsign - $last_equal_pos;
        @$lines[$linenumber] =~ s/=/' ' x abs($diff) .'='/e if ( $diff < 0 );
    }
}

sub get_max_equal_pos_from_lines {
    my $lines       = $_[0];
    my $linenumbers = $_[1];
    my @equalpositions;
    foreach my $linenumber (@$linenumbers) {
        push( @equalpositions, index( @$lines[$linenumber], '=' ) );
    }
    return max_from_array(@equalpositions);
}

sub max_from_array {
    my $max;
    foreach (@_) { $max = $_ if $_ > $max }
    return $max;
}

sub read_file {
    my ($file) = @_;
    local (*F);
    my $r;
    my (@r);
    open( F, "<$file" ) || die "ERROR: open $file: $!\n\n";

    #   binmode( F, ":crlf" );
    @r = <F>;
    close(F);
    return @r if wantarray;
    return join( "", @r );
}

sub write_file {
    my ( $f, @data ) = @_;
    local (*F);
    open( F, ">$f" ) || die "ERROR: open >$f: $!\n\n";

    #   binmode( F, ":crlf" );
    ( print F @data ) || die "ERROR: write $f: $!\n\n";
    close(F)          || die "ERROR: close $f: $!\n\n";
    return 1;
}

sub get_on_error_resume_next {
    my $text = $_[0];
    $$text =~ s/on\s+error\s+resume\s+next/%[resumenext]%/gis;
}

sub put_on_error_resume_next {
    my $text = $_[0];
    $$text =~ s/%\[resumenext\]%/On Error Resume Next/gis;
}

sub put_function_comments_to_declaration {
    my $text = $_[0];
    $$text =~ s/%\[comment\]%\n\nFunction/%[comment]%\nFunction/gis;
    $$text =~ s/%\[comment\]%\n\nSub/%[comment]%\nSub/gis;
}

sub singlelineifthen {
    my $line = $_[0];

    $$line =~ s/if.*?then\s+[^%]//gi;
}
