#!C:\strawberry\perl\bin\perl.exe
use strict;
use warnings;

use Cwd qw(getcwd);
use File::Copy qw(copy);
use Win32::OLE;
use Win32::OLE::Const;
use Time::Piece ();

use App::Rad;
App::Rad->run();

sub setup {
    my $c = shift;
    $c->register_commands(
        {
            honbun => 'create honbun.',
            mokuji => 'create mokuji. argument: --honbun=HONBUN_DATA',
            sakuin => 'create sakuin.',
            csp    => 'change sakuin page number start. arguments: --start=START_PAGE_NUMBER --indd=INDD_FILE',
            tsume  => 'create tsume index.',
            pdf    => 'create pdf. argument: --indd=INDD',
            push   => 'upload pdf files to the wke server. argument: --to=TO_DIRECOTRY',
            pull   => 'get data from the wke server. argument: --from=FROM_DIRECTORY',
        } );
}

sub honbun {
    my $files = getInitialFiles();
    $files->{tmp} .= "/template_honbun.indd";
    $files->{ind} .= "/honbun.indd";
    $files->{src} .= "/01_Honbun.jsx";
    copyFiles($files);
    doInDesignScript($files);
}

sub mokuji {
    my $c = shift;

    return 'You have to add honbun_data and page_data.' unless $c->options->{honbun};

    &makeMokujiData($c->options->{honbun});

    my $files = getInitialFiles();
    $files->{tmp} .= "/template_mokuji.indd";
    $files->{ind} .= "/mokuji.indd";
    $files->{src} .= "/02_Mokuji.jsx";
    copyFiles($files);
    doInDesignScript($files);
}

sub sakuin {
    my $files = getInitialFiles();
    $files->{tmp} .= "/template_mokuji.indd";
    $files->{ind} .= "/sakuin.indd";
    $files->{src} .= "/03_Sakuin.jsx";
    copyFiles($files);
    doInDesignScript($files);
}

sub csp {
    my $c = shift;
    if ( $c->options->{start} && $c->options->{indd} ) {
        my ($app, $wd) = &getInDesignObjects;
        $app->Open( getcwd . "/" . $c->options->{indd} );
        my $doc = $app->{ActiveDocument};
        $doc->{sections}->firstitem->{ContinueNumbering} = 0;
        $doc->{sections}->firstitem->{PageNumberStart} = $c->options->{start} + 0;
        $doc->Close( $wd->{idYes} );
    }
}

sub tsume {
    my $files = getInitialFiles();
    $files->{tmp} .= "/template_honbun.indd";
    $files->{ind} .= "/tsumeIndex.indd";
    $files->{src} .= "/04_TsumeIndex.jsx";
    copyFiles($files);
    doInDesignScript($files);
}

sub pdf {
    my $c = shift;
    if ($c->options->{indd}) {
        doPdf(getcwd() . "/" . $c->options->{indd});
    } else {
        return 'you have to add a .indd file.';
    }
}

sub push {
    my $c = shift;
    my $toserver = $c->options->{to} || '';

    if ( $toserver eq '' ) {
        open FH, 'config.conf';
        while (<FH>) {
            chomp;
            my @line = split ',', $_;
            if ( $line[0] eq 'toserver' ) {
                $toserver = $line[1];
            }
        }
    }

    exit if $toserver eq '';

    print "Ready to copy files on '$toserver' directory.\nAre you OK? (y/n)\n> ";
    $_ = <STDIN>;
    chomp;
    exit unless $_ eq 'y';

    my $dir = getcwd;
    opendir DIR, $dir;
    my @files = readdir DIR;
    close DIR;
    my %outfiles = (
        'honbun' => '',
        'mokuji' => '',
        'sakuin' => '',
        'page'   => 'pageNum.txt',
    );

    foreach my $file (@files) {
        next if $file eq ".";
        next if $file eq "..";
        if ( $file =~ m/^honbun.*pdf$/ ) {
            if ( $outfiles{honbun} eq '' || -M $outfiles{honbun} > -M $file ) {
                $outfiles{honbun} = $file;
            }
        } elsif ( $file =~ m/^mokuji.*pdf$/ ) {
            if ( $outfiles{mokuji} eq '' || -M $outfiles{mokuji} > -M $file ) {
                $outfiles{mokuji} = $file;
            }
        } elsif ( $file =~ m/^sakuin.*pdf$/ ) {
            if ( $outfiles{sakuin} eq '' || -M $outfiles{sakuin} > -M $file ) {
                $outfiles{sakuin} = $file;
            }
        }
    }
    foreach my $key ( keys %outfiles ) {
        if ( -e $outfiles{$key} ) {
            copy $outfiles{$key}, "$toserver\\$outfiles{$key}" if $outfiles{$key} ne '';
            print "Pushed $outfiles{$key}\n";
        }
    }
}

sub pull {
    my $c = shift;
    my $fromserver = $c->options->{from} || '';
    if ( $fromserver eq '' ) {
        open FH, 'config.conf';
        while (<FH>) {
            chomp;
            my @line = split ',', $_;
            if ( $line[0] eq 'fromserver' ) {
                $fromserver = $line[1];
            }
        }
    }

    exit if $fromserver eq '';

    my $curdir = getcwd;
    opendir DIR, $fromserver;
    my @files = readdir DIR;
    close DIR;
    my %getfiles = ('honbun' => '',
                    'sakuin' => '');
    foreach my $file ( @files ) {
        next if $file eq '.';
        next if $file eq '..';
        if ( $file =~ m/^Iyaku.*txt$/ ) {
            if ( $getfiles{honbun} eq '' || -M $getfiles{honbun} > -M $file ) {
                $getfiles{honbun} = $file;
            }
        } elsif ( $file =~ m/^索引.*txt$/ ) {
            if ( $getfiles{sakuin} eq '' || -M $getfiles{sakuin} > -M $file ) {
                $getfiles{sakuin} = $file;
            }
        }
    }
    foreach my $key ( keys %getfiles ) {
        if ( $getfiles{$key} ne '' ) {
            copy "$fromserver\\$getfiles{$key}", $curdir if $getfiles{$key} ne '';
            print "Pulled $getfiles{$key}\n";
        }
    }
}

sub getInitialFiles {
    print "Now setting.\n";
    my $path = getcwd();
    {'tmp' => $path, 'ind' => $path, 'src' => $path};
}

sub copyFiles {
    my $files = shift;
    copy $files->{tmp}, $files->{ind};
}

sub getInDesignObjects {
    return (
            Win32::OLE->new("InDesign.Application.CS5_J"),
            Win32::OLE::Const->Load("Adobe InDesign CS5_J"),
            );
}

sub doInDesignScript {
    my $files = shift;;

    print "Now making.\n";
    my ($app, $wd) = &getInDesignObjects;
    $app->DoScript($files->{src},
                   $wd->{idJavascript},
                   [],
                   $wd->{idFastEntireScript},
                   '');
    if ( $app->ActiveDocument ) {
        print "Now saving.\n";
        $app->ActiveDocument->close($wd->{idYes},
                                    &getUniqueFileName($files->{ind}));
    }
}

sub getUniqueFileName {
    my $name = substr($_[0], 0, rindex($_[0], ".indd"));
    my $buf = &makeFileName($name);
    for (my $i = 1; -e $buf; $i++) {
        $buf = &makeFileName($name);
    }
    $buf;
}

sub makeFileName {
    my $t = Time::Piece::localtime();
    my $filename = $_[0] || "foobar";
    $filename . "_" . $t->ymd("") . "_" . $t->hms("") . ".indd";
}

sub makeMokujiData {
    my $filename = shift;
    open HF, "<", $filename;
    open PF, "<", "pageNum.txt";
    open OUT, ">", "mokujiData.txt";

    my @honbun = <HF>;
    my @page   = <PF>;

    shift @honbun;

    my $len = @honbun;
    for (my $i = 0; $i < $len; $i++) {
        chomp $honbun[$i];
        chomp $page[$i];
        print OUT $honbun[$i] . "\t" . $page[$i] . "\n";
    }
}

sub doPdf {
    my $filename = shift;
    
    my ($app, $wd) = &getInDesignObjects;

    $app->Open($filename) or die "Couldn't open $filename.\n";
    my $doc = $app->{ActiveDocument};
    my $den = \%{$app->{PDFExportPresets}->{"Denshoku_ver"}};

    $den->{IncludeBookmarks}        = 1; # ブックマーク
    $den->{ExportReaderSpreads}     = 1; # 見開き印刷

    $den->{CropMarks}               = 1; # 内トンボ
    $den->{BleedMarks}              = 1; # 外トンボ
    $den->{RegistrationMarks}       = 1; # センタートンボ
    $den->{ColorBars}               = 0; # カラーバー
    $den->{PageInformationMarks}    = 0; # ページ情報

    $den->{UseDocumentBleedWithPDF} = 1; # ドキュメントの裁ち落とし設定を使用
    $den->{IncludeSlugWithPDF}      = 1; # 印刷可能領域を含む

    $filename =~ s/indd$/pdf/;
    $doc->AsynchronousExportFile($wd->{idPDFType},
                                 $filename,
                                 1,
                                 $den,
        );
    $doc->Close;
}
