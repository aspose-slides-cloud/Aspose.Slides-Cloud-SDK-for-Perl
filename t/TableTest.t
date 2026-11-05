=begin comment

Copyright (c) 2019 Aspose Pty Ltd

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

=end comment

=cut

BEGIN {push @INC, "../lib"}

use File::Slurp;
use Archive::Zip;
use JSON;

use AsposeSlidesCloud::TestUtils;
use AsposeSlidesCloud::SlidesApi;
use AsposeSlidesCloud::SlidesAsyncApi;
use AsposeSlidesCloud::Object::TableCell;
use AsposeSlidesCloud::Object::TableRow;
use AsposeSlidesCloud::Object::Paragraph;
use AsposeSlidesCloud::Object::Portion;
use AsposeSlidesCloud::Object::TableCellMergeOptions;

use Test::More;
use Test::Exception;

use strict;
use warnings;

my $utils = AsposeSlidesCloud::TestUtils->new();

subtest 'update table cell' => sub {
    eval {
		my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);

        my $dto = AsposeSlidesCloud::Object::TableCell->new();
        $dto->{text} = "Test text";

		my %params = (
            'name' => "test.pptx", 
            'password' => "password", 
            'folder' => "TempSlidesSDK", 
            'slide_index' => 9,
            'shape_index' => 1, 
            'row_index' => 2, 
            'cell_index' => 1, 
            'dto' => $dto);
        my $response = $utils->{testSlidesApi}->update_table_cell(%params);
        is($response->{text}, "Test text");
	};
    if ($@) {
        fail("update_table_cell raised an exception: $@");
    }
    pass();
};

subtest 'create table row' => sub {
    eval {
		my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);

        my $cell0 = AsposeSlidesCloud::Object::TableCell->new();
        $cell0->{text} = "Cell 1";
        my $cell1 = AsposeSlidesCloud::Object::TableCell->new();
        $cell1->{text} = "Cell 2";
        my $cell2 = AsposeSlidesCloud::Object::TableCell->new();
        $cell2->{text} = "Cell 3";
        my $cell3 = AsposeSlidesCloud::Object::TableCell->new();
        $cell3->{text} = "Cell 4";
        my @cells = ( $cell0, $cell1, $cell2, $cell3 );

        my $dto = AsposeSlidesCloud::Object::TableRow->new();
        $dto->{minimal_height} = 30;
        $dto->{cells} = \@cells;

		my %params = (
            'name' => "test.pptx", 
            'password' => "password", 
            'folder' => "TempSlidesSDK", 
            'slide_index' => 9,
            'shape_index' => 1, 
            'dto' => $dto);
        my $response = $utils->{testSlidesApi}->create_table_row(%params);
        is(scalar @{$response->{cells}}, scalar @{$dto->{cells}});
        is($response->{minimal_height}, $dto->{minimal_height});
	};
    if ($@) {
        fail("create_table_row raised an exception: $@");
    }
    pass();
};

subtest 'delete table row' => sub {
    eval {
		my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);

		my %params = (
            'name' => "test.pptx", 
            'password' => "password", 
            'folder' => "TempSlidesSDK", 
            'slide_index' => 9,
            'shape_index' => 1, 
            'row_index' => 2,
            'with_attached_rows' => 'true');
        my $response = $utils->{testSlidesApi}->delete_table_row(%params);

        is(scalar @{$response->{rows}}, 2);
	};
    if ($@) {
        fail("delete_table_row raised an exception: $@");
    }
    pass();
};

subtest 'update table row' => sub {
    eval {
		my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);

        my $dto = AsposeSlidesCloud::Object::TableRow->new();
        $dto->{minimal_height} = 30;

		my %params = (
            'name' => "test.pptx", 
            'password' => "password", 
            'folder' => "TempSlidesSDK", 
            'slide_index' => 9,
            'shape_index' => 1, 
            'row_index' => 1,
            'dto' => $dto);

        my $response = $utils->{testSlidesApi}->update_table_row(%params);

        is($response->{minimal_height}, $dto->{minimal_height});
	};
    if ($@) {
        fail("update_table_row raised an exception: $@");
    }
    pass();
};

subtest 'merge table cells' => sub {
    eval {
		my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);

        my $dto = AsposeSlidesCloud::Object::TableCellMergeOptions->new();
        $dto->{first_row_index} = 1;
        $dto->{first_cell_index} = 1;
        $dto->{last_row_index} = 2;
        $dto->{last_cell_index} = 2;
        $dto->{allow_splitting} = 'true';

		my %params = (
            'name' => "test.pptx", 
            'password' => "password", 
            'folder' => "TempSlidesSDK", 
            'slide_index' => 9,
            'shape_index' => 1, 
            'table_cell_merge_options' => $dto);

        my $response = $utils->{testSlidesApi}->merge_table_cells(%params);

        is($response->{rows}[0]{cells}[0]->col_span, 2);
        is($response->{rows}[0]{cells}[0]->row_span, 2);
	};
    if ($@) {
        fail("merge_table_cells raised an exception: $@");
    }
    pass();
};

subtest 'split table cell by width' => sub {
    eval {
		my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);

		my %params = (
            'name' => "test.pptx", 
            'password' => "password", 
            'folder' => "TempSlidesSDK", 
            'split_type' => 'SplitByWidth',
            'value' => 10,
            'slide_index' => 9,
            'shape_index' => 1, 
            'row_index' => 1, 
            'cell_index' => 1);

        my $response = $utils->{testSlidesApi}->split_table_cell(%params);

        is(scalar @{$response->{rows}[0]{cells}}, 5);
	};
    if ($@) {
        fail("split_table_cell raised an exception: $@");
    }
    pass();
};

subtest 'split table cell by height' => sub {
    eval {
		my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);

		my %params = (
            'name' => "test.pptx", 
            'password' => "password", 
            'folder' => "TempSlidesSDK", 
            'split_type' => 'SplitByHeight',
            'value' => 10,
            'slide_index' => 9,
            'shape_index' => 1, 
            'row_index' => 1, 
            'cell_index' => 1);

        my $response = $utils->{testSlidesApi}->split_table_cell(%params);

        is(scalar @{$response->{rows}}, 5);
	};
    if ($@) {
        fail("split_table_cell raised an exception: $@");
    }
    pass();
};

subtest 'split table cell by col span' => sub {
    eval {
		my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);

		my %params = (
            'name' => "test.pptx", 
            'password' => "password", 
            'folder' => "TempSlidesSDK", 
            'split_type' => 'SplitByColSpan',
            'value' => 1,
            'slide_index' => 9,
            'shape_index' => 1, 
            'row_index' => 3, 
            'cell_index' => 1);

        my $response = $utils->{testSlidesApi}->split_table_cell(%params);

        is($response->{rows}[0]{cells}[0]->col_span, undef);
	};
    if ($@) {
        fail("split_table_cell raised an exception: $@");
    }
    pass();
};

subtest 'split table cell by row span' => sub {
    eval {
		my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);

		my %params = (
            'name' => "test.pptx", 
            'password' => "password", 
            'folder' => "TempSlidesSDK", 
            'split_type' => 'SplitByRowSpan',
            'value' => 1,
            'slide_index' => 9,
            'shape_index' => 1, 
            'row_index' => 2, 
            'cell_index' => 3);

        my $response = $utils->{testSlidesApi}->split_table_cell(%params);

        is($response->{rows}[1]{cells}[2]->col_span, undef);
	};
    if ($@) {
        fail("split_table_cell raised an exception: $@");
    }
    pass();
};

subtest 'get table cell paragraphs' => sub {
    eval {
		my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);

		my %params = (
            'name' => "test.pptx", 
            'password' => "password", 
            'folder' => "TempSlidesSDK", 
            'slide_index' => 9,
            'shape_index' => 1, 
            'row_index' => 1, 
            'cell_index' => 1);

        my $response = $utils->{testSlidesApi}->get_table_cell_paragraphs(%params);

        is(scalar @{$response->{paragraph_links}}, 1);
	};
    if ($@) {
        fail("get_table_cell_paragraphs raised an exception: $@");
    }
    pass();
};

subtest 'get table cell paragraph' => sub {
    eval {
		my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);

		my %params = (
            'name' => "test.pptx", 
            'password' => "password", 
            'folder' => "TempSlidesSDK", 
            'slide_index' => 9,
            'shape_index' => 1, 
            'row_index' => 1, 
            'cell_index' => 1,
            'paragraph_index' => 1);

        my $response = $utils->{testSlidesApi}->get_table_cell_paragraph(%params);

        is(scalar @{$response->{portion_list}}, 2);
	};
    if ($@) {
        fail("get_table_cell_paragraph raised an exception: $@");
    }
    pass();
};

subtest 'create table cell paragraph' => sub {
    eval {
		my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);

        my $portion0 = AsposeSlidesCloud::Object::Portion->new();
        $portion0->{text} = "Portion 1";
        my $portion1 = AsposeSlidesCloud::Object::Portion->new();
        $portion1->{text} = "Portion 2";
        my @portions = ($portion0, $portion1);
        my $dto = AsposeSlidesCloud::Object::Paragraph->new();
        $dto->{portion_list} = \@portions;

		my %params = (
            'name' => "test.pptx", 
            'password' => "password", 
            'folder' => "TempSlidesSDK", 
            'slide_index' => 9,
            'shape_index' => 1, 
            'row_index' => 1, 
            'cell_index' => 1,
            'dto' => $dto);

        my $response = $utils->{testSlidesApi}->create_table_cell_paragraph(%params);

        is(scalar @{$response->{portion_list}}, 2);
	};
    if ($@) {
        fail("create_table_cell_paragraph raised an exception: $@");
    }
    pass();
};

subtest 'update table cell paragraph' => sub {
    eval {
		my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);

        my $portion0 = AsposeSlidesCloud::Object::Portion->new();
        $portion0->{text} = "Portion 1";
        my $portion1 = AsposeSlidesCloud::Object::Portion->new();
        $portion1->{text} = "Portion 2";
        my @portions = ($portion0, $portion1);
        my $dto = AsposeSlidesCloud::Object::Paragraph->new();
        $dto->{portion_list} = \@portions;

		my %params = (
            'name' => "test.pptx", 
            'password' => "password", 
            'folder' => "TempSlidesSDK", 
            'slide_index' => 9,
            'shape_index' => 1, 
            'row_index' => 1, 
            'cell_index' => 1,
            'paragraph_index' => 1,
            'dto' => $dto);

        my $response = $utils->{testSlidesApi}->update_table_cell_paragraph(%params);

        is(scalar @{$response->{portion_list}}, 2);
	};
    if ($@) {
        fail("update_table_cell_paragraph raised an exception: $@");
    }
    pass();
};

subtest 'delete table cell paragraph' => sub {
    eval {
		my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);

		my %params = (
            'name' => "test.pptx", 
            'password' => "password", 
            'folder' => "TempSlidesSDK", 
            'slide_index' => 9,
            'shape_index' => 1, 
            'row_index' => 1, 
            'cell_index' => 1,
            'paragraph_index' => 1);

        my $response = $utils->{testSlidesApi}->delete_table_cell_paragraph(%params);

        is(scalar @{$response->{paragraph_links}}, 0);
	};
    if ($@) {
        fail("delete_table_cell_paragraph raised an exception: $@");
    }
    pass();
};

subtest 'download table cell paragraph' => sub {
    eval {
        my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);

        my %params = (
            'name' => "test.pptx", 
            'password' => "password", 
            'folder' => "TempSlidesSDK", 
            'slide_index' => 9,
            'shape_index' => 1, 
            'row_index' => 1, 
            'cell_index' => 1,
            'paragraph_index' => 1,
            'format' => 'Png');

        my $response = $utils->{testSlidesApi}->download_table_cell_paragraph(%params);
        ok(length($response) > 0);
    };
    if ($@) {
        fail("download_table_cell_paragraph raised an exception: $@");
    }
    pass();
};

subtest 'save table cell paragraph' => sub {
    eval {
        my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);

        my $out_path = "TempSlidesSDK/tableCellParagraph.png";
        my %params = (
            'name' => "test.pptx", 
            'password' => "password", 
            'folder' => "TempSlidesSDK", 
            'slide_index' => 9,
            'shape_index' => 1, 
            'row_index' => 1, 
            'cell_index' => 1,
            'paragraph_index' => 1,
            'format' => 'Png',
            'out_path' => $out_path);

        $utils->{testSlidesApi}->save_table_cell_paragraph(%params);

        my %exists_params = ('path' => $out_path);
        my $exists = $utils->{testSlidesApi}->object_exists(%exists_params);
        ok($exists->{exists});
    };
    if ($@) {
        fail("save_table_cell_paragraph raised an exception: $@");
    }
    pass();
};

subtest 'get table cell portions' => sub {
    eval {
		my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);

		my %params = (
            'name' => "test.pptx", 
            'password' => "password", 
            'folder' => "TempSlidesSDK", 
            'slide_index' => 9,
            'shape_index' => 1, 
            'row_index' => 1, 
            'cell_index' => 1,
            'paragraph_index' => 1);

        my $response = $utils->{testSlidesApi}->get_table_cell_portions(%params);

        is(scalar @{$response->{items}}, 2);
	};
    if ($@) {
        fail("get_table_cell_portions raised an exception: $@");
    }
    pass();
};

subtest 'get table cell portion' => sub {
    eval {
		my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);

		my %params = (
            'name' => "test.pptx", 
            'password' => "password", 
            'folder' => "TempSlidesSDK", 
            'slide_index' => 9,
            'shape_index' => 1, 
            'row_index' => 1, 
            'cell_index' => 1,
            'paragraph_index' => 1,
            'portion_index' => 1);

        my $response = $utils->{testSlidesApi}->get_table_cell_portion(%params);

        is($response->{text}, "Header");
	};
    if ($@) {
        fail("get_table_cell_portion raised an exception: $@");
    }
    pass();
};

subtest 'create table cell portion' => sub {
    eval {
		my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);

        my $dto = AsposeSlidesCloud::Object::Portion->new();
        $dto->{text} = "Portion 1";

		my %params = (
            'name' => "test.pptx", 
            'password' => "password", 
            'folder' => "TempSlidesSDK", 
            'slide_index' => 9,
            'shape_index' => 1, 
            'row_index' => 1, 
            'cell_index' => 1,
            'paragraph_index' => 1,
            'dto' => $dto);

        my $response = $utils->{testSlidesApi}->create_table_cell_portion(%params);

        is($response->{text}, $dto->{text});
	};
    if ($@) {
        fail("create_table_cell_portion raised an exception: $@");
    }
    pass();
};

subtest 'update table cell portion' => sub {
    eval {
		my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);

        my $dto = AsposeSlidesCloud::Object::Portion->new();
        $dto->{text} = "Portion 1";

		my %params = (
            'name' => "test.pptx", 
            'password' => "password", 
            'folder' => "TempSlidesSDK", 
            'slide_index' => 9,
            'shape_index' => 1, 
            'row_index' => 1, 
            'cell_index' => 1,
            'paragraph_index' => 1,
            'portion_index' => 1,
            'dto' => $dto);

        my $response = $utils->{testSlidesApi}->update_table_cell_portion(%params);

        is($response->{text}, $dto->{text});
	};
    if ($@) {
        fail("update_table_cell_portion raised an exception: $@");
    }
    pass();
};

subtest 'delete table cell portion' => sub {
    eval {
		my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);

		my %params = (
            'name' => "test.pptx", 
            'password' => "password", 
            'folder' => "TempSlidesSDK", 
            'slide_index' => 9,
            'shape_index' => 1, 
            'row_index' => 1, 
            'cell_index' => 1,
            'paragraph_index' => 1,
            'portion_index' => 1);

        my $response = $utils->{testSlidesApi}->delete_table_cell_portion(%params);

        is(scalar @{$response->{items}}, 1);
	};
    if ($@) {
        fail("delete_table_cell_portion raised an exception: $@");
    }
    pass();
};

subtest 'import table from workbook' => sub {
    eval {
        my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);

        my $source = read_file("TestData/oleObject.xlsx", { binmode => ':raw' });
        my %params = (
            'name' => "test.pptx",
            'slide_index' => 9,
            'worksheet_name' => "Sheet1",
            'cell_range' => "A1:B5",
            'document' => $source,
            'password' => "password",
            'folder' => "TempSlidesSDK");
        my $result = $utils->{testSlidesApi}->import_table_from_workbook(%params);
        is($result->{type}, "Table");
    };
    if ($@) {
        fail("import_table_from_workbook raised an exception: $@");
    }
    pass();
};

subtest 'import table from workbook by path' => sub {
    eval {
        my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);

        my $source = read_file("TestData/oleObject.xlsx", { binmode => ':raw' });
        my %upload_params = ('path' => "TempSlidesSDK/oleObject.xlsx", 'file' => $source);
        $utils->{testSlidesApi}->upload_file(%upload_params);

        my %params = (
            'name' => "test.pptx",
            'slide_index' => 9,
            'worksheet_name' => "Sheet1",
            'cell_range' => "A1:B5",
            'workbook_path' => "TempSlidesSDK/oleObject.xlsx",
            'password' => "password",
            'folder' => "TempSlidesSDK");
        my $result = $utils->{testSlidesApi}->import_table_from_workbook(%params);
        is($result->{type}, "Table");
    };
    if ($@) {
        fail("import_table_from_workbook raised an exception: $@");
    }
    pass();
};

done_testing;