package Serge::Engine::Plugin::parse_xlsx;

use parent Serge::Engine::Plugin::Base::Parser;
use parent Serge::Interface::PluginHost;

use Serge::Util qw(generate_hash);

use Spreadsheet::Read;
use Spreadsheet::Write;

my $START_ROW_IDX           = 2;
my $EMPTY_FIELD_PLACEHOLDER = 'sfsjhdfgslfkjsl';

sub name {
    return 'XLS/XLSX parser plugin';
}

sub init {
    my $self = shift;

    $self->SUPER::init(@_);

    $self->merge_schema({
        source_column    => 'STRING',
        target_column    => 'STRING',
        context_column   => 'STRING'
    });
}

sub read_file {
    my ($self, $filename, $calc_hash) = @_;
    
    print "\n:: parse_xlsx::read_file($filename, $calc_hash)\n";

    my $output = "";

    my $book   = ReadData($filename);
    my $sheet  = $book->[1];
    my $maxRow = $sheet->{maxrow};

    print "\n\tContext column: $self->{data}->{context_column}";
    print "\n\tSource column:  $self->{data}->{source_column}";
    print "\n\tTarget column:  $self->{data}->{target_column}";
    print "\n\tNumber of rows: $maxRow\n\n";

    my $rowIdx = $START_ROW_IDX;

    while ($rowIdx <= $maxRow) {
        my $contextCell = $self->{data}->{context_column} . $rowIdx;
        my $sourceCell  = $self->{data}->{source_column} . $rowIdx;

        my ($cc, $cr)   = cell2cr($contextCell);
        my ($sc, $sr)   = cell2cr($sourceCell);

        my $context     = $sheet->{cell}[$cc][$cr];
        my $source      = $sheet->{cell}[$sc][$sr];

        if ($context eq '') {
            $context = $EMPTY_FIELD_PLACEHOLDER;
        }

        if ($source eq '') {
            $source = $EMPTY_FIELD_PLACEHOLDER;
        }

        $output .= "$context\n";
        $output .= "$source\n";

        $rowIdx++;
    }
  
    my $hash = generate_hash($output) if $calc_hash;

    return ($output, $hash);
}

sub serialize {
    my ($self, $strref) = @_;
    print "\n:: parse_xlsx::serialize just returns the input and hash\n\n";
    return ($$strref, generate_hash($$strref));
}

sub write_file {
    my ($self, $source_file, $target_file, $strref) = @_;
    print "\n:: parse_xlsx::write_file($target_file)\n";

    my @targets;

    my $key   = undef;
    my $value = undef;

    foreach my $line (split(/\n/, $$strref)) {
        if ($key eq undef) {
            $key = $line;
        } else {
            $value = $line;
            push(@targets, $value);
            $key = undef;
        }
    }

    my $targetCell               = $self->{data}->{target_column} . $START_ROW_IDX;
    my ($targetColumnIdx, undef) = cell2cr($targetCell);
    $targetColumnIdx            -= 1; # convert to 0-based from 1-based;

    my $book = ReadData($source_file);
    my $out  = Spreadsheet::Write->new(file => $target_file);

    my $sheet  = $book->[1];
    my $maxRow = $sheet->{maxrow};

    my $numTargets = @targets;
    my $numRows    = $maxRow - 1;
    die "Number of target strings $numTargets != $numRows rows in the source file" if ($numTargets != $numRows);

    my $rowIdx = 1; # starting from the first row to include column labels

    while ($rowIdx <= $maxRow) {
        my @row = Spreadsheet::Read::row($sheet, $rowIdx);
        
        if ($rowIdx > 1 ) {
            my $targetIdx = $rowIdx - 2; # -2 shift as row indexes are 1-based and the first row is not in the targets
            if (not @targets[$targetIdx] eq $EMPTY_FIELD_PLACEHOLDER) {
                @row[$targetColumnIdx] = @targets[$targetIdx];
            }
        }

        $out->addrow(@row);

        $rowIdx++;
    }

    $out->close();
}

sub parse {
    my ($self, $source_text, $callbackref, $lang) = @_;
    print "\n:: parse_xlsx::parse\n";

    print "\n\tlang = $lang\n\n" if $self->{parent}->{debug};

    die 'callbackref not specified' unless $callbackref;

    my @output;

    my $context = undef;
    my $value   = undef;
    my $key     = 0;

    foreach my $line (split(/\n/, $$source_text)) {
        if ($context eq undef) {
            $context = $line;
            $key++;
            push @output, $context;
        } else {
            my $translated_str;
            $value = $line;

            if ($context eq $EMPTY_FIELD_PLACEHOLDER) {
                $context = $key;
            }

            if ($value eq $EMPTY_FIELD_PLACEHOLDER) {
                $translated_str = $value;
            } else {
                $translated_str = &$callbackref(
                    $value,
                    $context,
                    undef, # ?????  we don't pass hints, they just go to the output ??????????
                    undef,
                    $lang,
                    $key
                );
            }

            if ($lang) {
                $value = $translated_str;
            }

            push @output, $value;

            $context = undef;
            $value   = undef;
        }
    }

    my $result = join("\n", @output);

    return $result;
}

1;
