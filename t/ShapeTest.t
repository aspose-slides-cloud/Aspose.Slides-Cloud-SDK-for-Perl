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

use Test::More;
use Test::Exception;

use AsposeSlidesCloud::TestUtils;
use AsposeSlidesCloud::SlidesApi;
use AsposeSlidesCloud::SlidesAsyncApi;
use AsposeSlidesCloud::Object::Shape;
use AsposeSlidesCloud::Object::GraphicalObject;
use AsposeSlidesCloud::Object::PictureFrame;
use AsposeSlidesCloud::Object::AudioFrame;
use AsposeSlidesCloud::Object::VideoFrame;
use AsposeSlidesCloud::Object::OleObjectFrame;
use AsposeSlidesCloud::Object::SmartArt;
use AsposeSlidesCloud::Object::Chart;
use AsposeSlidesCloud::Object::Table;
use AsposeSlidesCloud::Object::GroupShape;
use AsposeSlidesCloud::Object::Connector;
use AsposeSlidesCloud::Object::GeometryPaths;
use AsposeSlidesCloud::Object::MoveToPathSegment;
use AsposeSlidesCloud::Object::LineToPathSegment;
use AsposeSlidesCloud::Object::ClosePathSegment;
use AsposeSlidesCloud::Object::Portion;
use AsposeSlidesCloud::Object::GradientFill;
use AsposeSlidesCloud::Object::GradientFillStop;
use AsposeSlidesCloud::Object::ZoomFrame;
use AsposeSlidesCloud::Object::SectionZoomFrame;
use AsposeSlidesCloud::Object::OleObjectFrame;

use strict;
use warnings;
use MIME::Base64 qw(encode_base64);

my $utils = AsposeSlidesCloud::TestUtils->new();

subtest 'get shapes' => sub {
    eval {
      my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
      $utils->{testSlidesApi}->copy_file(%copy_params);

      my %params = ('name' => "test.pptx", 'slide_index' => 3, 'password' => "password", 'folder' => "TempSlidesSDK");
      my $shapes = $utils->{testSlidesApi}->get_shapes(%params);
      is(scalar @{$shapes->{shapes_links}}, 2);
    };
    if ($@) {
        fail("get_shapes raised an exception: $@");
    }
    pass();
};

subtest 'get shapes by type' => sub {
    eval {
      my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
      $utils->{testSlidesApi}->copy_file(%copy_params);

      my %params = (
        'name' => "test.pptx", 
        'slide_index' => 3, 
        'password' => "password", 
        'folder' => "TempSlidesSDK",
        'shape_type' => 'Chart');
      my $shapes = $utils->{testSlidesApi}->get_shapes(%params);
      is(scalar @{$shapes->{shapes_links}}, 2);
    };
    if ($@) {
        fail("get_shapes raised an exception: $@");
    }
    pass();
};

subtest 'get sub-shapes' => sub {
    eval {
      my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
      $utils->{testSlidesApi}->copy_file(%copy_params);

      my %params = (
        'name' => "test.pptx", 
        'slide_index' => 1, 
        'password' => "password", 
        'folder' => "TempSlidesSDK",
        'sub_shape' => '4');
      my $shapes = $utils->{testSlidesApi}->get_shapes(%params);
      is(scalar @{$shapes->{shapes_links}}, 2);
    };
    if ($@) {
        fail("get_shapes raised an exception: $@");
    }
    pass();
};

subtest 'get shape' => sub {
    eval {
      my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
      $utils->{testSlidesApi}->copy_file(%copy_params);

      my %params = ('name' => "test.pptx", 'slide_index' => 3, 'shape_index' => 1, 'password' => "password", 'folder' => "TempSlidesSDK");
      my $shape = $utils->{testSlidesApi}->get_shape(%params);
      is($shape->{type}, 'Chart');
    };
    if ($@) {
        fail("get_shape raised an exception: $@");
    }
    pass();
};

subtest 'get sub-shape' => sub {
    eval {
      my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
      $utils->{testSlidesApi}->copy_file(%copy_params);

      my %params = (
        'name' => "test.pptx", 
        'slide_index' => 1, 
        'shape_index' => 4, 
        'password' => "password", 
        'folder' => "TempSlidesSDK",
        'sub_shape' => '1');
      my $shape = $utils->{testSlidesApi}->get_shape(%params);
      is($shape->{type}, 'Shape');
    };
    if ($@) {
        fail("get_shape raised an exception: $@");
    }
    pass();
};

subtest 'shape load save' => sub {
    eval {
      my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
      $utils->{testSlidesApi}->copy_file(%copy_params);

      my %params = ('name' => "test.pptx", 'slide_index' => 3, 'shape_index' => 1, 'password' => "password", 'folder' => "TempSlidesSDK");
      my $dto = $utils->{testSlidesApi}->get_shape(%params);
      $params{dto} = $dto;
      my $shape = $utils->{testSlidesApi}->update_shape(%params);
      is(ref $shape, "AsposeSlidesCloud::Object::Chart");
    };
    if ($@) {
        fail("update_shape raised an exception: $@");
    }
    pass();
};

subtest 'shape add' => sub {
    eval {
      my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
      $utils->{testSlidesApi}->copy_file(%copy_params);

      my $dto = AsposeSlidesCloud::Object::Shape->new();
      $dto->{shape_type} = "Callout1";
      my %params = ('name' => "test.pptx", 'slide_index' => 1, 'dto' => $dto, 'password' => "password", 'folder' => "TempSlidesSDK");
      my $shape = $utils->{testSlidesApi}->create_shape(%params);
      is(ref $shape, "AsposeSlidesCloud::Object::Shape");
    };
    if ($@) {
        fail("create_shape raised an exception: $@");
    }
    pass();
};

subtest 'shape empty' => sub {
    eval {
      my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
      $utils->{testSlidesApi}->copy_file(%copy_params);

      my $dto = AsposeSlidesCloud::Object::Shape->new();
      my %params = ('name' => "test.pptx", 'slide_index' => 1, 'dto' => $dto, 'password' => "password", 'folder' => "TempSlidesSDK");
      $utils->{testSlidesApi}->create_shape(%params);
    };
    if ($@) {
        if ($@ =~ m/API Exception\((\d+)\): (.*) /s) {
            is($1, 400);
        } else {
            fail("Invalid exception type: $@");
        }
    } else {
        fail("Shape with undefined type should not have been created");
    }
    pass();
};

subtest 'graphical object empty' => sub {
    eval {
      my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
      $utils->{testSlidesApi}->copy_file(%copy_params);

      my $dto = AsposeSlidesCloud::Object::GraphicalObject->new();
      my %params = ('name' => "test.pptx", 'slide_index' => 1, 'dto' => $dto, 'password' => "password", 'folder' => "TempSlidesSDK");
      $utils->{testSlidesApi}->create_shape(%params);
    };
    if ($@) {
        if ($@ =~ m/API Exception\((\d+)\): (.*) /s) {
            is($1, 400);
        } else {
            fail("Invalid exception type: $@");
        }
    } else {
        fail("GraphicalObject should not have been created");
    }
    pass();
};

subtest 'picture frame add' => sub {
    eval {
      my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
      $utils->{testSlidesApi}->copy_file(%copy_params);

      my $dto = AsposeSlidesCloud::Object::PictureFrame->new();
      my $fill = AsposeSlidesCloud::Object::PictureFill->new();
      $fill->{base64_data} = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsQAAA7EAZUrDhsAAAANSURBVBhXY5g+ffp/AAZTAsWGL27gAAAAAElFTkSuQmCC";
      $dto->{picture_fill_format} = $fill;
      my %params = ('name' => "test.pptx", 'slide_index' => 1, 'dto' => $dto, 'password' => "password", 'folder' => "TempSlidesSDK");
      my $shape = $utils->{testSlidesApi}->create_shape(%params);
      is(ref $shape, "AsposeSlidesCloud::Object::PictureFrame");
    };
    if ($@) {
        fail("create_shape raised an exception: $@");
    }
    pass();
};

subtest 'picture frame empty' => sub {
    eval {
      my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
      $utils->{testSlidesApi}->copy_file(%copy_params);

      my $dto = AsposeSlidesCloud::Object::PictureFrame->new();
      my %params = ('name' => "test.pptx", 'slide_index' => 1, 'dto' => $dto, 'password' => "password", 'folder' => "TempSlidesSDK");
      $utils->{testSlidesApi}->create_shape(%params);
    };
    if ($@) {
        if ($@ =~ m/API Exception\((\d+)\): (.*) /s) {
            is($1, 400);
        } else {
            fail("Invalid exception type: $@");
        }
    } else {
        fail("PictureFrame with undefined data should not have been created");
    }
    pass();
};

subtest 'audio frame add' => sub {
    eval {
      my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
      $utils->{testSlidesApi}->copy_file(%copy_params);

      my $dto = AsposeSlidesCloud::Object::AudioFrame->new();
      $dto->{base64_data} = "bXAzc2FtcGxl";
      my %params = ('name' => "test.pptx", 'slide_index' => 1, 'dto' => $dto, 'password' => "password", 'folder' => "TempSlidesSDK");
      my $shape = $utils->{testSlidesApi}->create_shape(%params);
      is(ref $shape, "AsposeSlidesCloud::Object::AudioFrame");
    };
    if ($@) {
        fail("create_shape raised an exception: $@");
    }
    pass();
};

subtest 'audio frame empty' => sub {
    eval {
      my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
      $utils->{testSlidesApi}->copy_file(%copy_params);

      my $dto = AsposeSlidesCloud::Object::AudioFrame->new();
      my %params = ('name' => "test.pptx", 'slide_index' => 1, 'dto' => $dto, 'password' => "password", 'folder' => "TempSlidesSDK");
      $utils->{testSlidesApi}->create_shape(%params);
    };
    if ($@) {
        if ($@ =~ m/API Exception\((\d+)\): (.*) /s) {
            is($1, 400);
        } else {
            fail("Invalid exception type: $@");
        }
    } else {
        fail("AudioFrame with undefined data should not have been created");
    }
    pass();
};

subtest 'video frame add' => sub {
    eval {
      my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
      $utils->{testSlidesApi}->copy_file(%copy_params);

      my $dto = AsposeSlidesCloud::Object::VideoFrame->new();
      $dto->{base64_data} = "bXAzc2FtcGxl";
      my %params = ('name' => "test.pptx", 'slide_index' => 1, 'dto' => $dto, 'password' => "password", 'folder' => "TempSlidesSDK");
      my $shape = $utils->{testSlidesApi}->create_shape(%params);
      is(ref $shape, "AsposeSlidesCloud::Object::VideoFrame");
    };
    if ($@) {
        fail("create_shape raised an exception: $@");
    }
    pass();
};

subtest 'video frame empty' => sub {
    eval {
      my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
      $utils->{testSlidesApi}->copy_file(%copy_params);

      my $dto = AsposeSlidesCloud::Object::VideoFrame->new();
      my %params = ('name' => "test.pptx", 'slide_index' => 1, 'dto' => $dto, 'password' => "password", 'folder' => "TempSlidesSDK");
      $utils->{testSlidesApi}->create_shape(%params);
    };
    if ($@) {
        if ($@ =~ m/API Exception\((\d+)\): (.*) /s) {
            is($1, 400);
        } else {
            fail("Invalid exception type: $@");
        }
    } else {
        fail("VideoFrame with undefined data should not have been created");
    }
    pass();
};

subtest 'ole object frame empty' => sub {
    eval {
      my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
      $utils->{testSlidesApi}->copy_file(%copy_params);

      my $dto = AsposeSlidesCloud::Object::OleObjectFrame->new();
      my %params = ('name' => "test.pptx", 'slide_index' => 1, 'dto' => $dto, 'password' => "password", 'folder' => "TempSlidesSDK");
      $utils->{testSlidesApi}->create_shape(%params);
    };
    if ($@) {
        if ($@ =~ m/API Exception\((\d+)\): (.*) /s) {
            is($1, 400);
        } else {
            fail("Invalid exception type: $@");
        }
    } else {
        fail("OleObjectFrame should not have been created");
    }
    pass();
};

subtest 'smart art add' => sub {
    eval {
      my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
      $utils->{testSlidesApi}->copy_file(%copy_params);

      my $dto = AsposeSlidesCloud::Object::SmartArt->new();
      $dto->{x} = 0;
      $dto->{y} = 0;
      $dto->{width} = 300;
      $dto->{height} = 200;
      $dto->{layout} = 'BasicProcess';
      $dto->{quick_style} = 'SimpleFill';
      $dto->{color_style} = 'ColoredFillAccent1';
      my $node1 = AsposeSlidesCloud::Object::SmartArtNode->new();
      $node1->{text} = "First";
      $node1->{org_chart_layout} = 'Initial';
      my $sub_node1 = AsposeSlidesCloud::Object::SmartArtNode->new();
      $sub_node1->{text} = "SubFirst";
      $sub_node1->{org_chart_layout} = 'Initial';
      my @sub_nodes = ( $sub_node1 );
      $node1->{nodes} = \@sub_nodes;
      my $node2 = AsposeSlidesCloud::Object::SmartArtNode->new();
      $node2->{text} = "Second";
      $node2->{org_chart_layout} = 'Initial';
      my @nodes = ( $node1, $node2 );
      $dto->{nodes} = \@nodes;
      my %params = ('name' => "test.pptx", 'slide_index' => 1, 'dto' => $dto, 'password' => "password", 'folder' => "TempSlidesSDK");
      my $shape = $utils->{testSlidesApi}->create_shape(%params);
      is(ref $shape, "AsposeSlidesCloud::Object::SmartArt");
    };
    if ($@) {
        fail("create_shape raised an exception: $@");
    }
    pass();
};

subtest 'smart art add text formatting' => sub {
    eval {
      my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
      $utils->{testSlidesApi}->copy_file(%copy_params);

      my $portion = AsposeSlidesCloud::Object::Portion->new();
      $portion->{text} = 'New text';
      $portion->{font_height} = 24;
      $portion->{font_bold} = 1;
      $portion->{spacing} = 3;
      my $fill_format = AsposeSlidesCloud::Object::SolidFill->new();
      $fill_format->{color} = '#FFFFFF00';
      $portion->{fill_format} = $fill_format;

      my %params = (
        'name' => "test.pptx", 
        'slide_index' => 7, 
        'shape_index' => 1,
        'paragraph_index' => 1,
        'portion_index' => 1,
        'dto' => $portion, 
        'password' => "password", 
        'folder' => "TempSlidesSDK",
        'sub_shape' => '1/nodes/2');

      my $response = $utils->{testSlidesApi}->update_portion(%params);
      is($portion->{text}, $response->{text});
      is($portion->{font_height}, $response->{font_height});
      is($portion->{font_bold}, '1');
      is($portion->{spacing}, $response->{spacing});
      is ($portion->{fill_format}->{color}, $portion->{fill_format}->{color});

    };
    if ($@) {
        fail("update_portion raised an exception: $@");
    }
    pass();
};

subtest 'smart art empty' => sub {
    eval {
      my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
      $utils->{testSlidesApi}->copy_file(%copy_params);

      my $dto = AsposeSlidesCloud::Object::SmartArt->new();
      my %params = ('name' => "test.pptx", 'slide_index' => 1, 'dto' => $dto, 'password' => "password", 'folder' => "TempSlidesSDK");
      my $shape = $utils->{testSlidesApi}->create_shape(%params);
      is(ref $shape, "AsposeSlidesCloud::Object::SmartArt");
    };
    if ($@) {
        fail("create_shape raised an exception: $@");
    }
    pass();
};

subtest 'chart empty' => sub {
    eval {
      my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
      $utils->{testSlidesApi}->copy_file(%copy_params);

      my $dto = AsposeSlidesCloud::Object::Chart->new();
      my %params = ('name' => "test.pptx", 'slide_index' => 1, 'dto' => $dto, 'password' => "password", 'folder' => "TempSlidesSDK");
      $utils->{testSlidesApi}->create_shape(%params);
    };
    if ($@) {
        if ($@ =~ m/API Exception\((\d+)\): (.*) /s) {
            is($1, 500);
        } else {
            fail("Invalid exception type: $@");
        }
    } else {
        fail("Chart with undefinined series data should not have been created");
    }
    pass();
};

subtest 'table add' => sub {
    eval {
      my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
      $utils->{testSlidesApi}->copy_file(%copy_params);

      my $dto = AsposeSlidesCloud::Object::Table->new();
      $dto->{x} = 30;
      $dto->{y} = 20;
      $dto->{style} = 'MediumStyle2Accent1';
      my $row1 = AsposeSlidesCloud::Object::TableRow->new();
      my $cell11 = AsposeSlidesCloud::Object::TableCell->new();
      $cell11->{text} = "0.1";
      my $cell12 = AsposeSlidesCloud::Object::TableCell->new();
      $cell12->{text} = "0.2";
      my $cell13 = AsposeSlidesCloud::Object::TableCell->new();
      $cell13->{text} = "0.3";
      my $cell14 = AsposeSlidesCloud::Object::TableCell->new();
      $cell14->{text} = "0.4";
      my @cells1 = ( $cell11, $cell12, $cell13, $cell14 );
      $row1->{cells} = \@cells1;
      my $row2 = AsposeSlidesCloud::Object::TableRow->new();
      my $cell21 = AsposeSlidesCloud::Object::TableCell->new();
      $cell21->{text} = "1";
      my $cell22 = AsposeSlidesCloud::Object::TableCell->new();
      $cell22->{text} = "2-3";
      $cell22->{col_span} = 2;
      $cell22->{row_span} = 2;
      my $cell24 = AsposeSlidesCloud::Object::TableCell->new();
      $cell24->{text} = "4";
      my @cells2 = ( $cell21, $cell22, $cell24 );
      $row2->{cells} = \@cells2;
      my $row3 = AsposeSlidesCloud::Object::TableRow->new();
      my $cell31 = AsposeSlidesCloud::Object::TableCell->new();
      $cell31->{text} = "first";
      my $cell32 = AsposeSlidesCloud::Object::TableCell->new();
      $cell32->{text} = "last";
      my @cells3 = ( $cell31, $cell32 );
      $row3->{cells} = \@cells3;
      my $row4 = AsposeSlidesCloud::Object::TableRow->new();
      my $cell41 = AsposeSlidesCloud::Object::TableCell->new();
      $cell41->{text} = "3.1";
      my $cell42 = AsposeSlidesCloud::Object::TableCell->new();
      $cell42->{text} = "3.2";
      my $cell43 = AsposeSlidesCloud::Object::TableCell->new();
      $cell43->{text} = "3.3";
      my $cell44 = AsposeSlidesCloud::Object::TableCell->new();
      $cell44->{text} = "3.4";
      my @cells4 = ( $cell41, $cell42, $cell43, $cell44 );
      $row4->{cells} = \@cells4;
      my $row5 = AsposeSlidesCloud::Object::TableRow->new();
      my $cell51 = AsposeSlidesCloud::Object::TableCell->new();
      $cell51->{text} = "4.1";
      my $cell52 = AsposeSlidesCloud::Object::TableCell->new();
      $cell52->{text} = "4.2";
      my $cell53 = AsposeSlidesCloud::Object::TableCell->new();
      $cell53->{text} = "4.3";
      my $cell54 = AsposeSlidesCloud::Object::TableCell->new();
      $cell54->{text} = "4.4";
      my @cells5 = ( $cell51, $cell52, $cell53, $cell54 );
      $row5->{cells} = \@cells5;
      my @rows = ( $row1, $row2, $row3, $row4, $row5 );
      $dto->{rows} = \@rows;
      my $column1 = AsposeSlidesCloud::Object::TableColumn->new();
      $column1->{width} = 100;
      my $column2 = AsposeSlidesCloud::Object::TableColumn->new();
      $column2->{width} = 100;
      my $column3 = AsposeSlidesCloud::Object::TableColumn->new();
      $column3->{width} = 100;
      my $column4 = AsposeSlidesCloud::Object::TableColumn->new();
      $column4->{width} = 100;
      my @columns = ( $column1, $column2, $column3, $column4 );
      $dto->{columns} = \@columns;
      $dto->{first_row} = 1;
      $dto->{horizontal_banding} = 1;
      my %params = ('name' => "test.pptx", 'slide_index' => 1, 'dto' => $dto, 'password' => "password", 'folder' => "TempSlidesSDK");
      my $shape = $utils->{testSlidesApi}->create_shape(%params);
      is(ref $shape, "AsposeSlidesCloud::Object::Table");
    };
    if ($@) {
        fail("create_shape raised an exception: $@");
    }
    pass();
};

subtest 'table empty' => sub {
    eval {
      my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
      $utils->{testSlidesApi}->copy_file(%copy_params);

      my $dto = AsposeSlidesCloud::Object::Table->new();
      my %params = ('name' => "test.pptx", 'slide_index' => 1, 'dto' => $dto, 'password' => "password", 'folder' => "TempSlidesSDK");
      $utils->{testSlidesApi}->create_shape(%params);
    };
    if ($@) {
        if ($@ =~ m/API Exception\((\d+)\): (.*) /s) {
            is($1, 400);
        } else {
            fail("Invalid exception type: $@");
        }
    } else {
        fail("Table with undefinined cell data should not have been created");
    }
    pass();
};

subtest 'group shape empty' => sub {
    eval {
      my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
      $utils->{testSlidesApi}->copy_file(%copy_params);

      my $dto = AsposeSlidesCloud::Object::GroupShape->new();
      my %params = ('name' => "test.pptx", 'slide_index' => 1, 'dto' => $dto, 'password' => "password", 'folder' => "TempSlidesSDK");
      my $shape = $utils->{testSlidesApi}->create_shape(%params);
      is(ref $shape, "AsposeSlidesCloud::Object::GroupShape");
    };
    if ($@) {
        fail("create_shape raised an exception: $@");
    }
    pass();
};

subtest 'connector add' => sub {
    eval {
      my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
      $utils->{testSlidesApi}->copy_file(%copy_params);

      my $dto = AsposeSlidesCloud::Object::Connector->new();
      $dto->{shape_type} = 'BentConnector3';
      my $start_shape = AsposeSlidesCloud::Object::ResourceUri->new();
      $start_shape->{href} = "https://api.aspose.cloud/v3.0/slides/myPresentation.pptx/slides/1/shapes/1";
      $dto->{start_shape_connected_to} = $start_shape;
      my $end_shape = AsposeSlidesCloud::Object::ResourceUri->new();
      $end_shape->{href} = "https://api.aspose.cloud/v3.0/slides/myPresentation.pptx/slides/1/shapes/2";
      $dto->{end_shape_connected_to} = $end_shape;
      my %params = ('name' => "test.pptx", 'slide_index' => 1, 'dto' => $dto, 'password' => "password", 'folder' => "TempSlidesSDK");
      my $shape = $utils->{testSlidesApi}->create_shape(%params);
      is(ref $shape, "AsposeSlidesCloud::Object::Connector");
    };
    if ($@) {
        fail("create_shape raised an exception: $@");
    }
    pass();
};

subtest 'connector empty' => sub {
    eval {
      my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
      $utils->{testSlidesApi}->copy_file(%copy_params);

      my $dto = AsposeSlidesCloud::Object::Connector->new();
      my %params = ('name' => "test.pptx", 'slide_index' => 1, 'dto' => $dto, 'password' => "password", 'folder' => "TempSlidesSDK");
      my $shape = $utils->{testSlidesApi}->create_shape(%params);
      is(ref $shape, "AsposeSlidesCloud::Object::Connector");
    };
    if ($@) {
        fail("create_shape raised an exception: $@");
    }
    pass();
};

subtest 'create sub-shape' => sub {
    eval {
      my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
      $utils->{testSlidesApi}->copy_file(%copy_params);

      my $dto = AsposeSlidesCloud::Object::Shape->new();
      $dto->{shape_type} = 'Rectangle';
      $dto->{x} = 200;
      $dto->{y} = 200;
      $dto->{width} = 50;
      $dto->{height} = 50;

      my %params = (
        'name' => "test.pptx", 
        'slide_index' => 1, 
        'dto' => $dto, 
        'password' => "password", 
        'folder' => "TempSlidesSDK",
        'sub_shape' => '4');

      my $shape = $utils->{testSlidesApi}->create_shape(%params);
      is(ref $shape, "AsposeSlidesCloud::Object::Shape");
    };
    if ($@) {
        fail("create_shape raised an exception: $@");
    }
    pass();
};

subtest 'update shape' => sub {
    eval {
      my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
      $utils->{testSlidesApi}->copy_file(%copy_params);

      my $dto = AsposeSlidesCloud::Object::Shape->new();
      $dto->{width} = 200;
      $dto->{height} = 200;
      my $fill_format = AsposeSlidesCloud::Object::SolidFill->new();
      $fill_format->{color} = '#FFF5FF8A';
      $dto->{fill_format} = $fill_format;

      my %params = (
        'name' => "test.pptx", 
        'slide_index' => 1, 
        'shape_index' => 3,
        'dto' => $dto, 
        'password' => "password", 
        'folder' => "TempSlidesSDK");

      my $response = $utils->{testSlidesApi}->update_shape(%params);
      is($response->{width}, $dto->{width});
      is($response->{height}, $dto->{height});
      is($response->{fill_format}->{type}, 'Solid');
    };
    if ($@) {
        fail("update_shape raised an exception: $@");
    }
    pass();
};

subtest 'update sub-shape' => sub {
    eval {
      my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
      $utils->{testSlidesApi}->copy_file(%copy_params);

      my $dto = AsposeSlidesCloud::Object::Shape->new();
      $dto->{width} = 200;
      $dto->{height} = 200;
      my $fill_format = AsposeSlidesCloud::Object::GradientFill->new();
      $fill_format->{shape} = 'Linear';
      $fill_format->{direction} = 'FromCorner1';
      my $gradient_stop1 = AsposeSlidesCloud::Object::GradientFillStop->new();
      $gradient_stop1->{color} = '#FFF5FF8A';
      $gradient_stop1->{position} = 0;
      my $gradient_stop2 = AsposeSlidesCloud::Object::GradientFillStop->new();
      $gradient_stop2->{color} = '#FFF5FF8A';
      $gradient_stop2->{position} = 1;
      my @stops = ($gradient_stop1, $gradient_stop2);
      $fill_format->{stops} = \@stops;
      $dto->{fill_format} = $fill_format;


      my %params = (
        'name' => "test.pptx", 
        'slide_index' => 1, 
        'shape_index' => 4,
        'dto' => $dto, 
        'password' => "password", 
        'folder' => "TempSlidesSDK",
        'sub_shape' => '1');

      my $response = $utils->{testSlidesApi}->update_shape(%params);
      is($response->{width}, $dto->{width});
      is($response->{height}, $dto->{height});
      is($response->{fill_format}->{type}, 'Gradient');
    };
    if ($@) {
        fail("update_shape raised an exception: $@");
    }
    pass();
};

subtest 'delete shapes' => sub {
    eval {
      my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
      $utils->{testSlidesApi}->copy_file(%copy_params);

      my %params = ('name' => "test.pptx", 'slide_index' => 3, 'password' => "password", 'folder' => "TempSlidesSDK");
      my $response = $utils->{testSlidesApi}->delete_shapes(%params);
      is(scalar @{$response->{shapes_links}}, 0)
    };
    if ($@) {
        fail("delete_shapes raised an exception: $@");
    }
    pass();
};

subtest 'delete shapes indexes' => sub {
    eval {
      my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
      $utils->{testSlidesApi}->copy_file(%copy_params);

      my @indexes = (2);

      my %params = (
        'name' => "test.pptx", 
        'slide_index' => 3,
        'shapes' => \@indexes, 
        'password' => "password", 
        'folder' => "TempSlidesSDK");
      my $response = $utils->{testSlidesApi}->delete_shapes(%params);
      is(scalar @{$response->{shapes_links}}, 1)
    };
    if ($@) {
        fail("delete_shapes raised an exception: $@");
    }
    pass();
};

subtest 'delete sub-shapes' => sub {
    eval {
      my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
      $utils->{testSlidesApi}->copy_file(%copy_params);

      my %params = (
        'name' => "test.pptx", 
        'slide_index' => 1, 
        'password' => "password", 
        'folder' => "TempSlidesSDK",
        'sub_shape' => '4');
      my $response = $utils->{testSlidesApi}->delete_shapes(%params);
      is(scalar @{$response->{shapes_links}}, 0)
    };
    if ($@) {
        fail("delete_shapes raised an exception: $@");
    }
    pass();
};

subtest 'delete sub-shapes indexes' => sub {
    eval {
      my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
      $utils->{testSlidesApi}->copy_file(%copy_params);

      my @indexes = (2);

      my %params = (
        'name' => "test.pptx", 
        'slide_index' => 1, 
        'shapes' => \@indexes,
        'password' => "password", 
        'folder' => "TempSlidesSDK",
        'sub_shape' => '4');
      my $response = $utils->{testSlidesApi}->delete_shapes(%params);
      is(scalar @{$response->{shapes_links}}, 1)
    };
    if ($@) {
        fail("delete_shapes raised an exception: $@");
    }
    pass();
};

subtest 'delete shape' => sub {
    eval {
      my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
      $utils->{testSlidesApi}->copy_file(%copy_params);

      my %params = (
        'name' => "test.pptx", 
        'slide_index' => 1, 
        'shape_index' => 4,
        'password' => "password", 
        'folder' => "TempSlidesSDK");
      my $response = $utils->{testSlidesApi}->delete_shape(%params);
      is(scalar @{$response->{shapes_links}}, 3)
    };
    if ($@) {
        fail("delete_shape raised an exception: $@");
    }
    pass();
};

subtest 'delete sub-shape' => sub {
    eval {
      my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
      $utils->{testSlidesApi}->copy_file(%copy_params);

      my %params = (
        'name' => "test.pptx", 
        'slide_index' => 1, 
        'shape_index' => 4,
        'password' => "password", 
        'folder' => "TempSlidesSDK",
        'sub_shape' => '1');
      my $response = $utils->{testSlidesApi}->delete_shape(%params);
      is(scalar @{$response->{shapes_links}}, 1)
    };
    if ($@) {
        fail("delete_shape raised an exception: $@");
    }
    pass();
};

subtest 'shapes align' => sub {
    eval {
      my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
      $utils->{testSlidesApi}->copy_file(%copy_params);

      my $slide_index = 3;
      my $shape1_index = 1;
      my $shape2_index = 2;
      my $dto = AsposeSlidesCloud::Object::Connector->new();
      my %params1 = ('name' => "test.pptx", 'slide_index' => $slide_index, 'shape_index' => $shape1_index, 'password' => "password", 'folder' => "TempSlidesSDK");
      my $shape1 = $utils->{testSlidesApi}->get_shape(%params1);
      my %params2 = ('name' => "test.pptx", 'slide_index' => $slide_index, 'shape_index' => $shape2_index, 'password' => "password", 'folder' => "TempSlidesSDK");
      my $shape2 = $utils->{testSlidesApi}->get_shape(%params2);
      ok(abs($shape1->{x} - $shape2->{x}) > 1);

      my %params = ('name' => "test.pptx", 'slide_index' => $slide_index, 'alignment_type' => "AlignTop", 'password' => "password", 'folder' => "TempSlidesSDK");
      $utils->{testSlidesApi}->align_shapes(%params);
      $shape1 = $utils->{testSlidesApi}->get_shape(%params1);
      $shape2 = $utils->{testSlidesApi}->get_shape(%params2);
      ok(abs($shape1->{x} - $shape2->{x}) > 1);
      ok(abs($shape1->{y} - $shape2->{y}) < 1);

      my @shapes = ( 1, 2 );
      %params = ('name' => "test.pptx", 'slide_index' => $slide_index, 'alignment_type' => "AlignLeft", 'align_to_slide' => 1, 'shapes' => \@shapes, 'password' => "password", 'folder' => "TempSlidesSDK");
      $utils->{testSlidesApi}->align_shapes(%params);
      $shape1 = $utils->{testSlidesApi}->get_shape(%params1);
      $shape2 = $utils->{testSlidesApi}->get_shape(%params2);
      ok(abs($shape1->{x} - $shape2->{x}) < 1);
      ok(abs($shape1->{x}) < 1);
      ok(abs($shape1->{y} - $shape2->{y}) < 1);
    };
    if ($@) {
        fail("align_shapes raised an exception: $@");
    }
    pass();
};

subtest 'shapes align group' => sub {
    eval {
      my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
      $utils->{testSlidesApi}->copy_file(%copy_params);

      my $slide_index = 1;
      my $shape1_index = "1";
      my $shape2_index = "2";
      my $shape_index = 4;
      my $shape_index_str = "4";
      my $dto = AsposeSlidesCloud::Object::Connector->new();
      my %params1 = ('name' => "test.pptx", 'slide_index' => $slide_index, 'shape_index' => $shape_index, 'sub_shape' => $shape1_index, 'password' => "password", 'folder' => "TempSlidesSDK");
      my $shape1 = $utils->{testSlidesApi}->get_shape(%params1);
      my %params2 = ('name' => "test.pptx", 'slide_index' => $slide_index, 'shape_index' => $shape_index, 'sub_shape' => $shape2_index, 'password' => "password", 'folder' => "TempSlidesSDK");
      my $shape2 = $utils->{testSlidesApi}->get_shape(%params2);
      ok(abs($shape1->{x} - $shape2->{x}) > 1);

      my %params = ('name' => "test.pptx", 'slide_index' => $slide_index, 'alignment_type' => "AlignTop", 'sub_shape' => $shape_index_str, 'password' => "password", 'folder' => "TempSlidesSDK");
      $utils->{testSlidesApi}->align_shapes(%params);
      $shape1 = $utils->{testSlidesApi}->get_shape(%params1);
      $shape2 = $utils->{testSlidesApi}->get_shape(%params2);
      ok(abs($shape1->{x} - $shape2->{x}) > 1);
      ok(abs($shape1->{y} - $shape2->{y}) < 1);

      my @shapes = ( 1, 2 );
      %params = ('name' => "test.pptx", 'slide_index' => $slide_index, 'alignment_type' => "AlignLeft", 'sub_shape' => $shape_index_str, 'align_to_slide' => 1, 'shapes' => \@shapes, 'password' => "password", 'folder' => "TempSlidesSDK");
      $utils->{testSlidesApi}->align_shapes(%params);
      $shape1 = $utils->{testSlidesApi}->get_shape(%params1);
      $shape2 = $utils->{testSlidesApi}->get_shape(%params2);
      ok(abs($shape1->{x} - $shape2->{x}) < 1);
      ok(abs($shape1->{x}) < 1);
      ok(abs($shape1->{y} - $shape2->{y}) < 1);
    };
    if ($@) {
        fail("align_shapes raised an exception: $@");
    }
    pass();
};

subtest 'shapes geometry get' => sub {
    eval {
      my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
      $utils->{testSlidesApi}->copy_file(%copy_params);

      my %params = ('name' => "test.pptx", 'slide_index' => 4, 'shape_index' => 2, 'password' => "password", 'folder' => "TempSlidesSDK");
      my $paths = $utils->{testSlidesApi}->get_shape_geometry_path(%params);
      is(scalar @{$paths->{paths}}, 1);
    };
    if ($@) {
        fail("get_shape_geometry_path raised an exception: $@");
    }
    pass();
};


subtest 'shapes geometry set' => sub {
    eval {
      my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
      $utils->{testSlidesApi}->copy_file(%copy_params);

      my $dto = AsposeSlidesCloud::Object::GeometryPaths->new();
      my $path = AsposeSlidesCloud::Object::GeometryPath->new();
      my $start_segment = AsposeSlidesCloud::Object::MoveToPathSegment->new();
      $start_segment->{x} = 0;
      $start_segment->{y} = 0;
      my $line1 = AsposeSlidesCloud::Object::LineToPathSegment->new();
      $line1->{x} = 0;
      $line1->{y} = 200;
      my $line2 = AsposeSlidesCloud::Object::LineToPathSegment->new();
      $line2->{x} = 200;
      $line2->{y} = 300;
      my $line3 = AsposeSlidesCloud::Object::LineToPathSegment->new();
      $line3->{x} = 400;
      $line3->{y} = 200;
      my $line4 = AsposeSlidesCloud::Object::LineToPathSegment->new();
      $line4->{x} = 400;
      $line4->{y} = 0;
      my $end_segment = AsposeSlidesCloud::Object::ClosePathSegment->new();
      my @path_data = ( $start_segment, $line1, $line2, $line3, $line4, $end_segment );
      $path->{path_data} = \@path_data;
      my @path = ( $path );
      $dto->{paths} = \@path;
      my %params = ('name' => "test.pptx", 'slide_index' => 4, 'shape_index' => 1, 'dto' => $dto, 'password' => "password", 'folder' => "TempSlidesSDK");
      my $shape = $utils->{testSlidesApi}->set_shape_geometry_path(%params);
      ok($shape);
    };
    if ($@) {
        fail("set_shape_geometry_path raised an exception: $@");
    }
    pass();
};

subtest 'zoom frame add' => sub {
    eval {
      my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
      $utils->{testSlidesApi}->copy_file(%copy_params);

      my $dto = AsposeSlidesCloud::Object::ZoomFrame->new();
      $dto->{x} = 20;
      $dto->{y} = 20;
      $dto->{width} = 200;
      $dto->{height} = 100;
      $dto->{target_slide_index} = 2;

      my %params = (
        'name' => "test.pptx", 
        'slide_index' => 3,
        'dto' => $dto, 
        'password' => "password", 
        'folder' => "TempSlidesSDK");
      my $shape = $utils->{testSlidesApi}->create_shape(%params);
      is($shape->{type}, 'ZoomFrame');
      is($shape->{target_slide_index}, $dto->{target_slide_index});
    };
    if ($@) {
        fail("create_shape raised an exception: $@");
    }
    pass();
};

subtest 'section zoom frame add' => sub {
    eval {
      my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
      $utils->{testSlidesApi}->copy_file(%copy_params);

      my $dto = AsposeSlidesCloud::Object::SectionZoomFrame->new();
      $dto->{x} = 20;
      $dto->{y} = 20;
      $dto->{width} = 200;
      $dto->{height} = 100;
      $dto->{target_section_index} = 2;

      my %params = (
        'name' => "test.pptx", 
        'slide_index' => 3, 
        'dto' => $dto, 
        'password' => "password", 
        'folder' => "TempSlidesSDK");
      my $shape = $utils->{testSlidesApi}->create_shape(%params);
      is($shape->{type}, 'SectionZoomFrame');
      is($shape->{target_section_index}, $dto->{target_section_index});
    };
    if ($@) {
        fail("create_shape raised an exception: $@");
    }
    pass();
};

subtest 'ole object frame add by link' => sub {
    eval {
      my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
      $utils->{testSlidesApi}->copy_file(%copy_params);

      my $dto = AsposeSlidesCloud::Object::OleObjectFrame->new();
      $dto->{x} = 100;
      $dto->{y} = 100;
      $dto->{width} = 200;
      $dto->{height} = 200;
      $dto->{link_path} = 'oleObject.xlsx';
      $dto->{object_prog_id} = 'Excel.Sheet.8';

      my %params = (
        'name' => "test.pptx", 
        'slide_index' => 3, 
        'dto' => $dto, 
        'password' => "password", 
        'folder' => "TempSlidesSDK");

      my $shape = $utils->{testSlidesApi}->create_shape(%params);
      is($shape->{type}, 'OleObjectFrame');
      is($shape->{link_path}, $dto->{link_path});
    };
    if ($@) {
        fail("create_shape raised an exception: $@");
    }
    pass();
};

subtest 'ole object frame add embedded' => sub {
    eval {
      my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
      $utils->{testSlidesApi}->copy_file(%copy_params);

      my $source = read_file("TestData/oleObject.xlsx", { binmode => ':raw' });

      my $dto = AsposeSlidesCloud::Object::OleObjectFrame->new();
      $dto->{x} = 100;
      $dto->{y} = 100;
      $dto->{width} = 200;
      $dto->{height} = 200;
      $dto->{embedded_file_base64_data} = encode_base64($source);
      $dto->{embedded_file_extension} = 'xlsx';

      my %params = (
        'name' => "test.pptx", 
        'slide_index' => 3, 
        'dto' => $dto, 
        'password' => "password", 
        'folder' => "TempSlidesSDK");

      my $shape = $utils->{testSlidesApi}->create_shape(%params);
      is($shape->{type}, 'OleObjectFrame');
    };
    if ($@) {
        fail("create_shape raised an exception: $@");
    }
    pass();
};

subtest 'group shape add' => sub {
    eval {
      my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
      $utils->{testSlidesApi}->copy_file(%copy_params);

      my %params = ('name' => "test.pptx", 'slide_index' => 5, 'password' => "password", 'folder' => "TempSlidesSDK");
      my $shapes = $utils->{testSlidesApi}->get_shapes(%params);
      is(scalar @{$shapes->{shapes_links}}, 0);

      #add group shape
      my $dto = AsposeSlidesCloud::Object::GroupShape->new();
      %params = ('name' => "test.pptx", 'slide_index' => 5, 'dto' => $dto, 'password' => "password", 'folder' => "TempSlidesSDK");
      $utils->{testSlidesApi}->create_shape(%params);

      #add sub-shapes
      my $shape1 = AsposeSlidesCloud::Object::Shape->new();
      $shape1->{shape_type} = 'Rectangle';
      $shape1->{x} = 50;
      $shape1->{y} = 400;
      $shape1->{width} = 50;
      $shape1->{height} = 50;

      my $shape2 = AsposeSlidesCloud::Object::Shape->new();
      $shape2->{shape_type} = 'Ellipse';
      $shape2->{x} = 150;
      $shape2->{y} = 400;
      $shape2->{width} = 50;
      $shape2->{height} = 50;

      my $shape3 = AsposeSlidesCloud::Object::Shape->new();
      $shape3->{shape_type} = 'Triangle';
      $shape3->{x} = 250;
      $shape3->{y} = 400;
      $shape3->{width} = 50;
      $shape3->{height} = 50;

      %params = ('name' => "test.pptx", 'slide_index' => 5, 'dto' => $shape1, 'password' => "password", 'folder' => "TempSlidesSDK", 'sub_shape' => '1');
      $utils->{testSlidesApi}->create_shape(%params);
      %params = ('name' => "test.pptx", 'slide_index' => 5, 'dto' => $shape2, 'password' => "password", 'folder' => "TempSlidesSDK", 'sub_shape' => '1');
      $utils->{testSlidesApi}->create_shape(%params);
      %params = ('name' => "test.pptx", 'slide_index' => 5, 'dto' => $shape3, 'password' => "password", 'folder' => "TempSlidesSDK", 'sub_shape' => '1');
      $utils->{testSlidesApi}->create_shape(%params);

      %params = ('name' => "test.pptx", 'slide_index' => 5, 'password' => "password", 'folder' => "TempSlidesSDK");
      $shapes = $utils->{testSlidesApi}->get_shapes(%params);
      is(scalar @{$shapes->{shapes_links}}, 1);

      %params = ('name' => "test.pptx", 'slide_index' => 5, 'password' => "password", 'folder' => "TempSlidesSDK", 'sub_shape' => '1');
      $shapes = $utils->{testSlidesApi}->get_shapes(%params);
      is(scalar @{$shapes->{shapes_links}}, 3);
    };
    if ($@) {
        fail("create_shape raised an exception: $@");
    }
    pass();
};

subtest 'import shapes from svg' => sub {
    eval {
		my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);

        my $source = read_file("TestData/shapes.svg", { binmode => ':raw' });

        my @shapes_indexes = (1,3,5);

		my %params = (
            'name' => "test.pptx", 
            'slide_index' => 5, 
            'image' => $source,
            'x' => 50,
            'y' => 50,
            'width' => 300,
            'height' => 300,
            'shapes' => \@shapes_indexes,
            'group' => 0,
            'password' => "password", 
            'folder' => "TempSlidesSDK");
        my $shapes = $utils->{testSlidesApi}->import_shapes_from_svg(%params);
        is(scalar @{$shapes->{shapes_links}}, 3);

	};
    if ($@) {
        fail("import_shapes_from_svg raised an exception: $@");
    }
    pass();
};

subtest 'create smart art node' => sub {
    eval {
		my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);

		my %params = (
            'name' => "test.pptx", 
            'slide_index' => 7, 
            'smart_art_index' => 1,
            'text' => 'new root node',
            'password' => "password", 
            'folder' => "TempSlidesSDK");
        my $response = $utils->{testSlidesApi}->create_smart_art_node(%params);	
        is(scalar @{$response->{nodes}}, 2);
        is($response->{nodes}[1]->{text}, 'new root node');
	};
    if ($@) {
        fail("create_smart_art_node raised an exception: $@");
    }
    pass();
};

subtest 'create smart art sub-node' => sub {
    eval {
		my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);

		my %params = (
            'name' => "test.pptx", 
            'slide_index' => 7, 
            'smart_art_index' => 1,
            'sub_node'=> '1',
            'position' => '1',
            'text' => 'new sub-node',
            'password' => "password", 
            'folder' => "TempSlidesSDK");
        my $response = $utils->{testSlidesApi}->create_smart_art_node(%params);	
        is(scalar @{$response->{nodes}[0]{nodes}}, 5);
        is($response->{nodes}[0]{nodes}[0]->{text}, 'new sub-node');
	};
    if ($@) {
        fail("create_smart_art_node raised an exception: $@");
    }
    pass();
};

subtest 'create smart art sub-sub-node' => sub {
    eval {
		my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);

		my %params = (
            'name' => "test.pptx", 
            'slide_index' => 7, 
            'smart_art_index' => 1,
            'sub_node'=> '1/nodes/1',
            'text' => 'new sub-sub-node',
            'password' => "password", 
            'folder' => "TempSlidesSDK");
        my $response = $utils->{testSlidesApi}->create_smart_art_node(%params);	
        is(scalar @{$response->{nodes}[0]{nodes}[0]{nodes}}, 1);
        is($response->{nodes}[0]{nodes}[0]{nodes}[0]->{text}, 'new sub-sub-node');
	};
    if ($@) {
        fail("create_smart_art_node raised an exception: $@");
    }
    pass();
};

subtest 'delete smart art node' => sub {
    eval {
		my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);

		my %params = (
            'name' => "test.pptx", 
            'slide_index' => 7, 
            'smart_art_index' => 2,
            'node_index' => 1,
            'password' => "password", 
            'folder' => "TempSlidesSDK");
        my $response = $utils->{testSlidesApi}->delete_smart_art_node(%params);	
        is(scalar @{$response->{nodes}}, 2);
	};
    if ($@) {
        fail("delete_smart_art_node raised an exception: $@");
    }
    pass();
};

subtest 'delete smart art sub-node' => sub {
    eval {
		my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);

		my %params = (
            'name' => "test.pptx", 
            'slide_index' => 7, 
            'smart_art_index' => 1,
            'node_index' => 1,
            'sub_node' => 2,
            'password' => "password", 
            'folder' => "TempSlidesSDK");
        my $response = $utils->{testSlidesApi}->delete_smart_art_node(%params);	
        is(scalar @{$response->{nodes}[0]{nodes}}, 3);
	};
    if ($@) {
        fail("delete_smart_art_node raised an exception: $@");
    }
    pass();
};

subtest 'download shape from DTO' => sub {
    eval {


      my $dto = AsposeSlidesCloud::Object::Shape->new();
      $dto->{shape_type} = "Rectangle";
      $dto->{width} = 400;
      $dto->{height} = 200;
      my %params = ('format' => "png", 'dto' => $dto);
      my $response = $utils->{testSlidesApi}->download_shape_from_dto(%params);
      ok(length($response) != 0);
    };
    if ($@) {
        fail("download_shape_from_dto raised an exception: $@");
    }
    pass();
};

subtest 'smart art node default paragraph format' => sub {
    eval {
        my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);

        my %get_params = (
            'name' => "test.pptx",
            'slide_index' => 7,
            'shape_index' => 1,
            'password' => "password",
            'folder' => "TempSlidesSDK");
        my $existing = $utils->{testSlidesApi}->get_shape(%get_params);
        my $paragraph_format = AsposeSlidesCloud::Object::ParagraphFormat->new();
        my $bullet_fill = AsposeSlidesCloud::Object::SolidFill->new();
        $bullet_fill->{color} = "#FFFF0000";
        $paragraph_format->{bullet_fill_format} = $bullet_fill;
        $existing->{nodes}[0]{default_paragraph_format} = $paragraph_format;

        my %update_params = (
            'name' => "test.pptx",
            'slide_index' => 7,
            'shape_index' => 1,
            'dto' => $existing,
            'password' => "password",
            'folder' => "TempSlidesSDK");
        my $updated = $utils->{testSlidesApi}->update_shape(%update_params);
        ok(defined $updated->{nodes}[0]{default_paragraph_format});
        is($updated->{nodes}[0]{default_paragraph_format}{bullet_fill_format}{type}, "Solid");
        is($updated->{nodes}[0]{default_paragraph_format}{bullet_fill_format}{color}, "#FFFF0000");
    };
    if ($@) {
        fail("smart_art_node_default_paragraph_format raised an exception: $@");
    }
    pass();
};

subtest 'shape clone from slide' => sub {
    eval {
        my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);

        my %params = (
            'name' => "test.pptx",
            'slide_index' => 3,
            'shape_to_clone' => 1,
            'clone_from_slide' => 1,
            'password' => "password",
            'folder' => "TempSlidesSDK");
        my $result = $utils->{testSlidesApi}->create_shape(%params);
        ok(defined $result);
    };
    if ($@) {
        fail("shape_clone_from_slide raised an exception: $@");
    }
    pass();
};

done_testing;
