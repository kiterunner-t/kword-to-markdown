# Copyleft (C) KRT, 2014 by kiterunner_t
# ref: http://www.kiterunner.me/?p=427

use strict;
use warnings;

use Cwd;
use Encode;
use File::Basename;
use File::Find;
use File::Spec;
use File::stat;
use File::Temp qw/tempfile/;
use Getopt::Long;
use Win32::OLE qw(in);


sub filter_file;
sub dos_gbk_to_unix_utf8($$);
sub word_to_md($$$);
sub utf8_to_gbk($);
sub gbk_to_utf8($);
sub in_range(\@$);
sub chomp_text($);
sub print_head($$$$);
sub print_paragraph_text($$$$\@\@);
sub usage();


use constant {
    LEFT_INDENT => 21,
  };


if ($^O ne "MSWin32") {
  print "[ERROR] THE SCRIPT CAN BE RUNNING IN WINDOWS SYSTEM\n";
  exit 2;
}


my %opts = (
    "head-list-number"  => 1,
    "picture-format"    => "png",
    "picture-path"      => "images",
    "word-path"         => ".",
    "md-out-path"       => "../kiterunner.me/",
  );

Getopt::Long::Configure("bundling");
GetOptions(\%opts,
    "help|h",
    "verbose|v",
    "force|f",
    "picture-format=s",
    "head-list-number!",
    "gbk",
    "picture-path=s",
    "word-path=s",
    "md-out-file|o=s",
    "md-out-path=s",
  ) or die usage;

if (exists $opts{help}) {
  usage;
  exit 0;
}

my %styles = (
    Title         => "标题",
    Heading1      => "标题 1",
    Heading2      => "标题 2",
    Heading3      => "标题 3",
    Caption       => "题注",
    Emphasis      => "",
    ListParagraph => "列出段落",
    NormalObject  => "无间隔,program",
    text          => "正文",

    picture       => "图",
    table         => "表格",
  );

foreach my $k (keys %styles) {
  $styles{$k} = utf8_to_gbk $styles{$k};
}


my $word_app = Win32::OLE->GetActiveObject("Word.Application");
unless (defined $word_app) {
  $word_app = Win32::OLE->new("Word.Application", "Quit") or die "Couldn't run Word: $!";
}

my @files;


sub filter_file {
  my $name = $File::Find::name;
  $name =~ s#^\./##;
  push @files, $name if $name =~ m#\.docx$# && basename($name) !~ m/^~\$/;
}


if (@ARGV) {
  push @files, @ARGV;
} else {
  find \&filter_file, ".";
}

my $word_path = Cwd::realpath $opts{"word-path"};
my $out_path = Cwd::realpath $opts{"md-out-path"};

foreach my $file (@files) {
  my $file_withpath = File::Spec->catfile($word_path, $file);
  if (! -f $file_withpath) {
    print "    [WARNING] $file_withpath is not exist\n";
    next;
  }

  my $file_without_ext = (fileparse $file, "\.docx")[0];
  my $file_md = $file_without_ext . ".md";
  my $file_md_withpath = File::Spec->catfile($out_path, basename $file_md);

  if (exists $opts{force}) {
    print "[FORCE-CONVERTING] $file_withpath\n";

  } else {
    if (-f $file_md_withpath) {
      my $word_stat = stat $file_withpath or die "error: $!";
      my $md_stat = stat $file_md_withpath or die "error: $!";

      if ($word_stat->mtime > $md_stat->mtime) {
        print "  [RECONVERTING] $file_withpath\n";
      } else {
        print "    [INFO] $file_withpath is older than $file_md_withpath, ";
        print "use the --force option\n";
        next;
      }

    } else {
      print "[CONVERTING] $file_withpath\n";
    }
  }

  my $tmpfd = tempfile;
  word_to_md $word_app, $file_withpath, $tmpfd;

  if (! exists $opts{gbk} || $opts{gbk} == 0) {
    dos_gbk_to_unix_utf8 $file_md_withpath, $tmpfd;
  }

  close $tmpfd;
}


sub dos_gbk_to_unix_utf8($$) {
  my ($out_file, $tmpfd) = @_;

  open my $out_fd, "> :raw", $out_file or die "open $out_file error: $!";
  my $enc = find_encoding("gbk");
  my $enc_utf8 = find_encoding("utf-8");
  seek $tmpfd, 0, 0;
  while (<$tmpfd>) {
    s/[\r\n]+$//;
    print $out_fd $enc_utf8->encode($enc->decode($_)), "\n";
  }
  close $out_fd;
}


sub word_to_md($$$) {
  my ($word_app, $word_name_withpath, $tmpfd) = @_;

  my $close_word;
  my $doc;
  my $filename = basename $word_name_withpath;
  foreach my $d (in $word_app->Documents) {
    if ($d->Name eq $filename) {
      $doc = $d;
      last;
    }
  }

  unless (defined $doc) {
    $close_word = 1;
    $doc = $word_app->Documents->Open($word_name_withpath);
  }

  my @inline_shapes;
  foreach my $shape (in $doc->InlineShapes) {
    my $range = $shape->Range;
    my $s = {
        start => $range->Start,
        end => $range->End,
      };

    push @inline_shapes, $s;
  }

  my @tables;
  foreach my $table (in $doc->Tables) {
    my $range = $table->Range;
    my $t = {
        start => $range->Start,
        end => $range->End,
      };

    push @tables, $t;
  }

  my @links;
  foreach my $link (in $doc->Hyperlinks) {
    my $range = $link->Range;
    my $l = {
        start => $range->Start,
        end => $range->End,
        text => $range->Text,
        url => $link->Address,
      };

    $l->{text} =~ s/[\r\n]+$//;
    push @links, $l;
  }

  if (exists ${opts}{verbose}) {
    foreach my $table (@tables) {
      print "[debug] table start: ", $table->{start}, ", end: ", $table->{end}, "\n";
    }

    foreach my $link (@links) {
      print "[debug] link start: ", $link->{start}, ", end: ", $link->{end}, " ", $link->{text}, "\n";
    }
  }

  my $last_end = 0;
  my $ref_count = 0;
  my $begin_indent = 0;
  my $last_indent = 0;
  my $indent;
  my @refs;

  my @urls;

  my $head1 = 0;
  my $head2 = 0;
  my $head3 = 0;

  my $program = 0;

  foreach my $paragraph (in $doc->Paragraphs) {
    my $start = $paragraph->Range->start;
    my $end = $paragraph->Range->end;

    next if $start < $last_end;

    if ($last_end = in_range(@tables, $start)) {
      next;
    } elsif ($last_end = in_range(@inline_shapes, $start)) {
      next;
    }

    my $style = $paragraph->Format->Style->NameLocal;
    my $text = $paragraph->Range->Text;
    my $left_indent = int($paragraph->Format->LeftIndent);

    $text =~ s/[\r\n]+$//;

    if (exists $opts{verbose}) {
      print "[debug] style --> " . $paragraph->Format->Style->Type;
      print "[$start -> $end]-----$style--->>>($left_indent) $text\n";
    }

    if ($program > 0 && $style ne $styles{NormalObject}) {
      $program = 0;
      print $tmpfd "\n" if ! $text =~ m/\s*$/;
    }

    $last_indent = 0 if $style ne $styles{ListParagraph};

    if ($style eq $styles{Title}) {
      # ignore

    } elsif ($style eq $styles{Heading1}) {
      $head1++;
      $head2 = 0;
      $head3 = 0;
      print_head $tmpfd, 1, $head1, $text;

    } elsif ($style eq $styles{Heading2}) {
      $head2++;
      $head3 = 0;
      print_head $tmpfd, 2, "$head1.$head2", $text;

    } elsif ($style eq $styles{Heading3}) {
      $head3++;
      print_head $tmpfd, 3, "$head1.$head2.$head3", $text;

    } elsif ($style eq $styles{ListParagraph}) {
      if ($last_indent == 0) {
        print $tmpfd "\n";
        $begin_indent = $left_indent;
      } elsif ($last_indent < $left_indent) {
        print $tmpfd "\n";
      } elsif ($last_indent > $left_indent) {
        print $tmpfd "\n";
      }

      $indent = int(($left_indent - $begin_indent) / LEFT_INDENT) * 4;
      $last_indent = $left_indent;
      print $tmpfd " " x $indent . "* ";
      $ref_count = print_paragraph_text $tmpfd, $doc, $paragraph, $ref_count, @links, @urls;

    } elsif ($style eq $styles{NormalObject}) {
      $program++;
      print $tmpfd "\n" if $program == 1;
      print $tmpfd "    $text";
      # $ref_count = print_paragraph_text $tmpfd, $doc, $paragraph, $ref_count, @links, @urls;

    } elsif ($style eq $styles{Emphasis}) {
      print $tmpfd "**";
      $ref_count = print_paragraph_text $tmpfd, $doc, $paragraph, $ref_count, @links, @urls;
      print $tmpfd "**";

    } elsif ($style eq $styles{Caption}) {
      my $table = $styles{table};
      my $picture = $styles{picture};
      if ($text =~ /\s*(?:$table|$picture)\s*\d+\s+(.*)$/) {
        my $name = $1;
        $name =~ s/[\r\n]+//g;
        $ref_count++;
        push @refs, [$ref_count, $name];
        print $tmpfd "\n![$name][$ref_count]\n\n";
      }

    } else {
      $ref_count = print_paragraph_text $tmpfd, $doc, $paragraph, $ref_count, @links, @urls;
    }
  }

  $opts{"picture-path"} =~ s/\/*$/\// if $opts{"picture-path"} ne "";
  print $tmpfd "\n";
  foreach my $ref (@refs) {
    print $tmpfd "[", $ref->[0], "]: ", $opts{"picture-path"}, $ref->[1], ".", $opts{"picture-format"}, " \"", $ref->[1], "\"\n";
  }

  foreach my $url (@urls) {
    print $tmpfd "[", $url->[0], "]: ", $url->[1], "\n";
  }

  if ($close_word) {
    $doc->Close;
  }
}


sub utf8_to_gbk($) {
  my ($str) = @_;

  encode "gbk", decode("utf-8", $str);
}


sub gbk_to_utf8($) {
  my ($str) = @_;

  encode "utf-8", decode("gbk", $str);
}


sub in_range(\@$) {
  my ($ranges, $start) = @_;

  foreach my $range (@$ranges) {
    if ($range->{start} == $start) {
      return $range->{end};
    }
  }

  0;
}


sub chomp_text($) {
  my ($text) = @_;
  $text =~ s/[\r\n]+$/\n/;
  $text =~ s/</\\</g;
  $text;
}


sub print_head($$$$) {
  my ($fd, $count, $list_numbers, $text) = @_;

  print $fd "#";
  print $fd "#" x $count;
  if (exists $opts{"head-list-number"} && $opts{"head-list-number"} != 0) {
    print $fd " " . $list_numbers;
  }
  print $fd " " . $text . "\n";
}


sub print_paragraph_text($$$$\@\@) {
  my ($fd, $doc, $paragram, $refcount, $links, $urls) = @_;

  if (exists $opts{verbose}) {
    print "[debug] 1 ", $paragram->Range->Start, " ", $paragram->Range->End, "\n";
  }

  # If Range is [n, n+1], the method Range->Start will be fail
  if ($paragram->Range->End - $paragram->Range->Start == 1) {
    print $fd chomp_text($paragram->Range->Text);
    return $refcount;
  }

  my $text_range = $paragram->Range;
 
  foreach my $l (@$links) {
    if ($l->{start} >= $text_range->Start && $l->{start} < $text_range->End) {
      my $pre_link = $doc->Range($text_range->Start, $l->{start});

      if (exists $opts{verbose}) {
        print "[debug] 00 ", $pre_link->Start, " ", $pre_link->End, "\n";
      }

      $refcount++;
      push @$urls, [$refcount, $l->{url}];
      print $fd chomp_text($pre_link->Text), "[", $l->{text}, "][$refcount]";
      last if $l->{end} == $text_range->End;
      $text_range = $doc->Range($l->{end}, $text_range->End);
    }
  }

  print $fd chomp_text($text_range->Text);
  $refcount;
}


sub usage() {
  print <<"_EOF_";

Usage:
    perl word_to_md.pl [options] [xx.docx]

options:
    --help, -h
    --verbose, -v
    --force, -f         force the script to reproduce the markdown file
    --picture-format    "png" is the default
    --head-list-number  print head list number by default
    --gbk               the encoding of the markdown is ms-dos/gbk
                        the default is unix/utf-8
    --picture-path      the default path is "images"
    --word-path         the default word path is the current path
    --md-out-file, -o   the mardown file name
    --md-out-path

_EOF_
}


# @TODO write the test examples and codes
# @TODO move the repo to a stand alone project

__END__

# Change Log (MD)
* 2014/05/16 refactor codes and add the following features

    * implement converting word to markdown in batch mode, and recognize the markdown is  stable to the relative word file
    * refactor some codes in converting the words to markdown
    * add the option to allow converting the mardown file to UNIX/UTF-8 (default) or MS-DOS/GBK
    * add the optoin to produce title number (default) or not

* 2014/05/04 add the new features and fix bugs as follows

    * support (and test) office word 2013
    * support hyperlink
    * enhance head1-head3 for section numbers
    * fix listparagraph "first line" bug

* 2014/04/22 write the first version

    * support (and only test) office word 2007
    * support the features, suach as head1-head3, codes, listparagraph, pictures and tables

