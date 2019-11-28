#!/usr/bin/perl

use strict;
use warnings;
use utf8;

use Excel::Writer::XLSX;

# Create a new Excel workbook
my $workbook = Excel::Writer::XLSX->new('write03.xlsx');

#  Add and define a format
my $format = $workbook->add_format();    # Add a format
$format->set_font('ＭＳ Ｐゴシック');

{
    # Add a worksheet
    my $worksheet = $workbook->add_worksheet('セリーグ');

    # Write using row and column notation.
    my $col = my $row = 0;
    $worksheet->write( $row, $col, '2019年ペナントレース結果',
        $format );
    $worksheet->write( 1, $col, '順位', $format );
    $worksheet->write( 2, $col, '1位', $format );
    $worksheet->write( 3, $col, '2位', $format );
    $worksheet->write( 4, $col, '3位', $format );
    $worksheet->write( 5, $col, '4位', $format );
    $worksheet->write( 6, $col, '5位', $format );
    $worksheet->write( 7, $col, '6位', $format );

    # Write using A1 notation
    $worksheet->write( 'B2', 'チーム名', $format );
    $worksheet->write( 'B3', '巨　人', $format );
    $worksheet->write( 'B4', 'DeNA', $format );
    $worksheet->write( 'B5', '阪　神', $format );
    $worksheet->write( 'B6', '広　島', $format );
    $worksheet->write( 'B7', '中　日', $format );
    $worksheet->write( 'B8', 'ヤクルト', $format );

    $worksheet->write( 'C2', '試合', $format );
    $worksheet->write( 'C3', '143' );
    $worksheet->write( 'C4', '143' );
    $worksheet->write( 'C5', '143' );
    $worksheet->write( 'C6', '143' );
    $worksheet->write( 'C7', '143' );
    $worksheet->write( 'C8', '143' );

    $worksheet->write( 'D2', '勝利数', $format );
    $worksheet->write( 'D3', '77' );
    $worksheet->write( 'D4', '71' );
    $worksheet->write( 'D5', '69' );
    $worksheet->write( 'D6', '70' );
    $worksheet->write( 'D7', '68' );
    $worksheet->write( 'D8', '58' );

    $worksheet->write( 'E2', '敗北数', $format );
    $worksheet->write( 'E3', '64' );
    $worksheet->write( 'E4', '69' );
    $worksheet->write( 'E5', '68' );
    $worksheet->write( 'E6', '70' );
    $worksheet->write( 'E7', '73' );
    $worksheet->write( 'E8', '82' );

    $worksheet->write( 'F2', '引分数', $format );
    $worksheet->write( 'F3', '2' );
    $worksheet->write( 'F4', '3' );
    $worksheet->write( 'F5', '6' );
    $worksheet->write( 'F6', '3' );
    $worksheet->write( 'F7', '2' );
    $worksheet->write( 'F8', '2' );
}

{
    # Add a worksheet
    my $worksheet = $workbook->add_worksheet('パリーグ');

    # Write using row and column notation.
    my $col = my $row = 0;
    $worksheet->write( $row, $col, '2019年ペナントレース結果',
        $format );
    $worksheet->write( 1, $col, '順位', $format );
    $worksheet->write( 2, $col, '1位', $format );
    $worksheet->write( 3, $col, '2位', $format );
    $worksheet->write( 4, $col, '3位', $format );
    $worksheet->write( 5, $col, '4位', $format );
    $worksheet->write( 6, $col, '5位', $format );
    $worksheet->write( 7, $col, '6位', $format );

    # Write using A1 notation
    $worksheet->write( 'B2', 'チーム名', $format );
    $worksheet->write( 'B3', '西　武', $format );
    $worksheet->write( 'B4', 'ソフトバンク', $format );
    $worksheet->write( 'B5', '楽　天', $format );
    $worksheet->write( 'B6', 'ロッテ', $format );
    $worksheet->write( 'B7', '日本ハム', $format );
    $worksheet->write( 'B8', 'オリックス', $format );

    $worksheet->write( 'C2', '試合', $format );
    $worksheet->write( 'C3', '143' );
    $worksheet->write( 'C4', '143' );
    $worksheet->write( 'C5', '143' );
    $worksheet->write( 'C6', '143' );
    $worksheet->write( 'C7', '143' );
    $worksheet->write( 'C8', '143' );

    $worksheet->write( 'D2', '勝利数', $format );
    $worksheet->write( 'D3', '80' );
    $worksheet->write( 'D4', '76' );
    $worksheet->write( 'D5', '71' );
    $worksheet->write( 'D6', '69' );
    $worksheet->write( 'D7', '65' );
    $worksheet->write( 'D8', '61' );

    $worksheet->write( 'E2', '敗北数', $format );
    $worksheet->write( 'E3', '62' );
    $worksheet->write( 'E4', '62' );
    $worksheet->write( 'E5', '68' );
    $worksheet->write( 'E6', '70' );
    $worksheet->write( 'E7', '73' );
    $worksheet->write( 'E8', '75' );

    $worksheet->write( 'F2', '引分数', $format );
    $worksheet->write( 'F3', '1' );
    $worksheet->write( 'F4', '5' );
    $worksheet->write( 'F5', '4' );
    $worksheet->write( 'F6', '4' );
    $worksheet->write( 'F7', '5' );
    $worksheet->write( 'F8', '7' );

}

$workbook->close();
