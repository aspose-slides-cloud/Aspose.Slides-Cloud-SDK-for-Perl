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
use JSON;

use Test::More;
use Test::Exception;

use AsposeSlidesCloud::TestUtils;
use AsposeSlidesCloud::SlidesApi;
use AsposeSlidesCloud::SlidesAsyncApi;
use AsposeSlidesCloud::Object::Chart;
use AsposeSlidesCloud::Object::OneValueSeries;
use AsposeSlidesCloud::Object::Workbook;
use AsposeSlidesCloud::Object::Literals;
use AsposeSlidesCloud::Object::GradientFillStop;
use AsposeSlidesCloud::Object::GradientFill;
use AsposeSlidesCloud::Object::Legend;
use AsposeSlidesCloud::Object::ChartWall;
use AsposeSlidesCloud::Object::BlurEffect;
use AsposeSlidesCloud::Object::EffectFormat;

use strict;
use warnings;

my $utils = AsposeSlidesCloud::TestUtils->new();

subtest 'chart get' => sub {
    $utils->initialize('get_shape', '');
    eval {
        my %params = ('name' => 'test.pptx', 'slide_index' => 3, 'shape_index' => 1, 'password' => 'password', 'folder' => 'TempSlidesSDK');
        my $chart = $utils->{testSlidesApi}->get_shape(%params);
        is(scalar @{$chart->{series}}, 3);
        is(scalar @{$chart->{categories}}, 4);
    };
    if ($@) {
        fail("get_shape raised an exception: $@");
    }
    pass();
};

subtest 'chart create auto data source' => sub {
    $utils->initialize('get_shape', '');
    eval {
        my $dto = AsposeSlidesCloud::Object::Chart->new();
        $dto->{chart_type} = "ClusteredColumn";
        $dto->{width} = 400;
        $dto->{height} = 300;
        my $series1 = AsposeSlidesCloud::Object::OneValueSeries->new();
        $series1->{name} = "Series1";
        my $point11 = AsposeSlidesCloud::Object::OneValueChartDataPoint->new();
        $point11->{value} = 40;
        my $point12 = AsposeSlidesCloud::Object::OneValueChartDataPoint->new();
        $point12->{value} = 50;
        my $point13 = AsposeSlidesCloud::Object::OneValueChartDataPoint->new();
        $point13->{value} = 70;
        my @points1 = ( $point11, $point12, $point13 );
        $series1->{data_points} = \@points1;
        my $series2 = AsposeSlidesCloud::Object::OneValueSeries->new();
        $series2->{name} = "Series2";
        my $point21 = AsposeSlidesCloud::Object::OneValueChartDataPoint->new();
        $point21->{value} = 55;
        my $point22 = AsposeSlidesCloud::Object::OneValueChartDataPoint->new();
        $point22->{value} = 35;
        my $point23 = AsposeSlidesCloud::Object::OneValueChartDataPoint->new();
        $point23->{value} = 90;
        my @points2 = ( $point21, $point22, $point23 );
        $series2->{data_points} = \@points2;
        my @series = ( $series1, $series2 );
        $dto->{series} = \@series;
        my $category1 = AsposeSlidesCloud::Object::ChartCategory->new();
        $category1->{value} = "Category1";
        my $category2 = AsposeSlidesCloud::Object::ChartCategory->new();
        $category2->{value} = "Category2";
        my $category3 = AsposeSlidesCloud::Object::ChartCategory->new();
        $category3->{value} = "Category3";
        my @categories = ( $category1, $category2, $category3 );
        $dto->{categories} = \@categories;
        my %params = ('name' => 'test.pptx', 'slide_index' => 3, 'dto' => $dto, 'password' => 'password', 'folder' => 'TempSlidesSDK');
        my $chart = $utils->{testSlidesApi}->create_shape(%params);
        is(scalar @{$chart->{series}}, 2);
        is(scalar @{$chart->{categories}}, 3);
    };
    if ($@) {
        fail("get_shape raised an exception: $@");
    }
    pass();
};

subtest 'chart create workbook data source' => sub {
    $utils->initialize('get_shape', '');
    eval {
        my $dto = AsposeSlidesCloud::Object::Chart->new();
        $dto->{chart_type} = "ClusteredColumn";
        $dto->{width} = 400;
        $dto->{height} = 300;
        my $data_source_for_categories = AsposeSlidesCloud::Object::Workbook->new();
        $data_source_for_categories->{worksheet_index} = 1;
        $data_source_for_categories->{column_index} = 1;
        $data_source_for_categories->{row_index} = 2;
        $dto->{data_source_for_categories} = $data_source_for_categories;

        my $series1 = AsposeSlidesCloud::Object::OneValueSeries->new();
        my $data_source_for_series1_name = AsposeSlidesCloud::Object::Workbook->new();
        $data_source_for_series1_name->{worksheet_index} = 1;
        $data_source_for_series1_name->{column_index} = 2;
        $data_source_for_series1_name->{row_index} = 1;
        $series1->{data_source_for_series_name} = $data_source_for_series1_name;
        $series1->{name} = "Series1";
        my $data_source_for_series1_values = AsposeSlidesCloud::Object::Workbook->new();
        $data_source_for_series1_values->{worksheet_index} = 1;
        $data_source_for_series1_values->{column_index} = 2;
        $data_source_for_series1_values->{row_index} = 2;
        $series1->{data_source_for_values} = $data_source_for_series1_values;
        my $point11 = AsposeSlidesCloud::Object::OneValueChartDataPoint->new();
        $point11->{value} = 40;
        my $point12 = AsposeSlidesCloud::Object::OneValueChartDataPoint->new();
        $point12->{value} = 50;
        my $point13 = AsposeSlidesCloud::Object::OneValueChartDataPoint->new();
        $point13->{value} = 70;
        my @points1 = ( $point11, $point12, $point13 );
        $series1->{data_points} = \@points1;
        my $series2 = AsposeSlidesCloud::Object::OneValueSeries->new();
        my $data_source_for_series2_name = AsposeSlidesCloud::Object::Workbook->new();
        $data_source_for_series2_name->{worksheet_index} = 1;
        $data_source_for_series2_name->{column_index} = 3;
        $data_source_for_series2_name->{row_index} = 1;
        $series1->{data_source_for_series_name} = $data_source_for_series2_name;
        $series2->{name} = "Series2";
        my $data_source_for_series2_values = AsposeSlidesCloud::Object::Workbook->new();
        $data_source_for_series2_values->{worksheet_index} = 1;
        $data_source_for_series2_values->{column_index} = 3;
        $data_source_for_series2_values->{row_index} = 2;
        $series2->{data_source_for_values} = $data_source_for_series2_values;
        my $point21 = AsposeSlidesCloud::Object::OneValueChartDataPoint->new();
        $point21->{value} = 55;
        my $point22 = AsposeSlidesCloud::Object::OneValueChartDataPoint->new();
        $point22->{value} = 35;
        my $point23 = AsposeSlidesCloud::Object::OneValueChartDataPoint->new();
        $point23->{value} = 90;
        my @points2 = ( $point21, $point22, $point23 );
        $series2->{data_points} = \@points2;
        my @series = ( $series1, $series2 );
        $dto->{series} = \@series;
        my $category1 = AsposeSlidesCloud::Object::ChartCategory->new();
        $category1->{value} = "Category1";
        my $category2 = AsposeSlidesCloud::Object::ChartCategory->new();
        $category2->{value} = "Category2";
        my $category3 = AsposeSlidesCloud::Object::ChartCategory->new();
        $category3->{value} = "Category3";
        my @categories = ( $category1, $category2, $category3 );
        $dto->{categories} = \@categories;
        my %params = ('name' => 'test.pptx', 'slide_index' => 3, 'dto' => $dto, 'password' => 'password', 'folder' => 'TempSlidesSDK');
        my $chart = $utils->{testSlidesApi}->create_shape(%params);
        is(scalar @{$chart->{series}}, 2);
        is(scalar @{$chart->{categories}}, 3);
    };
    if ($@) {
        fail("get_shape raised an exception: $@");
    }
    pass();
};

subtest 'chart create literals data source' => sub {
    $utils->initialize('get_shape', '');
    eval {
        my $dto = AsposeSlidesCloud::Object::Chart->new();
        $dto->{chart_type} = "ClusteredColumn";
        $dto->{width} = 400;
        $dto->{height} = 300;
        my $data_source_for_categories = AsposeSlidesCloud::Object::Literals->new();
        $dto->{data_source_for_categories} = $data_source_for_categories;

        my $series1 = AsposeSlidesCloud::Object::OneValueSeries->new();
        my $data_source_for_series1_name = AsposeSlidesCloud::Object::Literals->new();
        $series1->{data_source_for_series_name} = $data_source_for_series1_name;
        $series1->{name} = "Series1";
        my $data_source_for_series1_values = AsposeSlidesCloud::Object::Literals->new();
        $series1->{data_source_for_values} = $data_source_for_series1_values;
        my $point11 = AsposeSlidesCloud::Object::OneValueChartDataPoint->new();
        $point11->{value} = 40;
        my $point12 = AsposeSlidesCloud::Object::OneValueChartDataPoint->new();
        $point12->{value} = 50;
        my $point13 = AsposeSlidesCloud::Object::OneValueChartDataPoint->new();
        $point13->{value} = 70;
        my @points1 = ( $point11, $point12, $point13 );
        $series1->{data_points} = \@points1;
        my $series2 = AsposeSlidesCloud::Object::OneValueSeries->new();
        my $data_source_for_series2_name = AsposeSlidesCloud::Object::Literals->new();
        $series1->{data_source_for_series_name} = $data_source_for_series2_name;
        $series2->{name} = "Series2";
        my $data_source_for_series2_values = AsposeSlidesCloud::Object::Literals->new();
        $series2->{data_source_for_values} = $data_source_for_series2_values;
        my $point21 = AsposeSlidesCloud::Object::OneValueChartDataPoint->new();
        $point21->{value} = 55;
        my $point22 = AsposeSlidesCloud::Object::OneValueChartDataPoint->new();
        $point22->{value} = 35;
        my $point23 = AsposeSlidesCloud::Object::OneValueChartDataPoint->new();
        $point23->{value} = 90;
        my @points2 = ( $point21, $point22, $point23 );
        $series2->{data_points} = \@points2;
        my @series = ( $series1, $series2 );
        $dto->{series} = \@series;
        my $category1 = AsposeSlidesCloud::Object::ChartCategory->new();
        $category1->{value} = "Category1";
        my $category2 = AsposeSlidesCloud::Object::ChartCategory->new();
        $category2->{value} = "Category2";
        my $category3 = AsposeSlidesCloud::Object::ChartCategory->new();
        $category3->{value} = "Category3";
        my @categories = ( $category1, $category2, $category3 );
        $dto->{categories} = \@categories;
        my %params = ('name' => 'test.pptx', 'slide_index' => 3, 'dto' => $dto, 'password' => 'password', 'folder' => 'TempSlidesSDK');
        my $chart = $utils->{testSlidesApi}->create_shape(%params);
        is(scalar @{$chart->{series}}, 2);
        is(scalar @{$chart->{categories}}, 3);
    };
    if ($@) {
        fail("get_shape raised an exception: $@");
    }
    pass();
};

subtest 'chart update' => sub {
    $utils->initialize('update_shape', '');
    eval {
        my $dto = AsposeSlidesCloud::Object::Chart->new();
        $dto->{chart_type} = "ClusteredColumn";
        $dto->{width} = 400;
        $dto->{height} = 300;
        my $series1 = AsposeSlidesCloud::Object::OneValueSeries->new();
        $series1->{name} = "Series1";
        my $point11 = AsposeSlidesCloud::Object::OneValueChartDataPoint->new();
        $point11->{value} = 40;
        my $point12 = AsposeSlidesCloud::Object::OneValueChartDataPoint->new();
        $point12->{value} = 50;
        my $point13 = AsposeSlidesCloud::Object::OneValueChartDataPoint->new();
        $point13->{value} = 70;
        my @points1 = ( $point11, $point12, $point13 );
        $series1->{data_points} = \@points1;
        my $series2 = AsposeSlidesCloud::Object::OneValueSeries->new();
        $series2->{name} = "Series2";
        my $point21 = AsposeSlidesCloud::Object::OneValueChartDataPoint->new();
        $point21->{value} = 55;
        my $point22 = AsposeSlidesCloud::Object::OneValueChartDataPoint->new();
        $point22->{value} = 35;
        my $point23 = AsposeSlidesCloud::Object::OneValueChartDataPoint->new();
        $point23->{value} = 90;
        my @points2 = ( $point21, $point22, $point23 );
        $series2->{data_points} = \@points2;
        my @series = ( $series1, $series2 );
        $dto->{series} = \@series;
        my $category1 = AsposeSlidesCloud::Object::ChartCategory->new();
        $category1->{value} = "Category1";
        my $category2 = AsposeSlidesCloud::Object::ChartCategory->new();
        $category2->{value} = "Category2";
        my $category3 = AsposeSlidesCloud::Object::ChartCategory->new();
        $category3->{value} = "Category3";
        my @categories = ( $category1, $category2, $category3 );
        $dto->{categories} = \@categories;
        my %params = ('name' => 'test.pptx', 'slide_index' => 3, 'shape_index' => 1, 'dto' => $dto, 'password' => 'password', 'folder' => 'TempSlidesSDK');
        my $chart = $utils->{testSlidesApi}->update_shape(%params);
        is(scalar @{$chart->{series}}, 2);
        is(scalar @{$chart->{categories}}, 3);
    };
    if ($@) {
        fail("update_shape raised an exception: $@");
    }
    pass();
};

subtest 'chart series create' => sub {
    $utils->initialize('create_chart_series', '');
    eval {
        my $dto = AsposeSlidesCloud::Object::OneValueSeries->new();
        $dto->{name} = "Series1";
        my $point1 = AsposeSlidesCloud::Object::OneValueChartDataPoint->new();
        $point1->{value} = 40;
        my $point2 = AsposeSlidesCloud::Object::OneValueChartDataPoint->new();
        $point2->{value} = 50;
        my $point3 = AsposeSlidesCloud::Object::OneValueChartDataPoint->new();
        $point3->{value} = 70;
        my @points = ( $point1, $point2, $point3 );
        $dto->{data_points} = \@points;
        my %params = ('name' => 'test.pptx', 'slide_index' => 3, 'shape_index' => 1, 'series' => $dto, 'password' => 'password', 'folder' => 'TempSlidesSDK');
        my $chart = $utils->{testSlidesApi}->create_chart_series(%params);
        is(scalar @{$chart->{series}}, 4);
        is(scalar @{$chart->{categories}}, 4);
    };
    if ($@) {
        fail("create_chart_series raised an exception: $@");
    }
    pass();
};

subtest 'chart series update' => sub {
    $utils->initialize('update_chart_series', '');
    eval {
        my $dto = AsposeSlidesCloud::Object::OneValueSeries->new();
        $dto->{name} = "Series3";
        my $point1 = AsposeSlidesCloud::Object::OneValueChartDataPoint->new();
        $point1->{value} = 40;
        my $point2 = AsposeSlidesCloud::Object::OneValueChartDataPoint->new();
        $point2->{value} = 50;
        my $point3 = AsposeSlidesCloud::Object::OneValueChartDataPoint->new();
        $point3->{value} = 14;
        my $point4 = AsposeSlidesCloud::Object::OneValueChartDataPoint->new();
        $point4->{value} = 70;
        my @points = ( $point1, $point2, $point3, $point4 );
        $dto->{data_points} = \@points;
        my %params = ('name' => 'test.pptx', 'slide_index' => 3, 'shape_index' => 1, 'series_index' => 2, 'series' => $dto, 'password' => 'password', 'folder' => 'TempSlidesSDK');
        my $chart = $utils->{testSlidesApi}->update_chart_series(%params);
        is(scalar @{$chart->{series}}, 3);
        is(scalar @{$chart->{categories}}, 4);
    };
    if ($@) {
        fail("update_chart_series raised an exception: $@");
    }
    pass();
};

subtest 'chart series delete' => sub {
    $utils->initialize('delete_chart_series', '');
    eval {
        my %params = ('name' => 'test.pptx', 'slide_index' => 3, 'shape_index' => 1, 'series_index' => 2, 'password' => 'password', 'folder' => 'TempSlidesSDK');
        my $chart = $utils->{testSlidesApi}->delete_chart_series(%params);
        is(scalar @{$chart->{series}}, 2);
        is(scalar @{$chart->{categories}}, 4);
    };
    if ($@) {
        fail("delete_chart_series raised an exception: $@");
    }
    pass();
};

subtest 'chart category create' => sub {
    $utils->initialize('create_chart_category', '');
    eval {
        my $dto = AsposeSlidesCloud::Object::ChartCategory->new();
        $dto->{value} = "New Category";
        my $point1 = AsposeSlidesCloud::Object::OneValueChartDataPoint->new();
        $point1->{value} = 40;
        my $point2 = AsposeSlidesCloud::Object::OneValueChartDataPoint->new();
        $point2->{value} = 50;
        my $point3 = AsposeSlidesCloud::Object::OneValueChartDataPoint->new();
        $point3->{value} = 14;
        my @points = ( $point1, $point2, $point3 );
        $dto->{data_points} = \@points;
        my %params = ('name' => 'test.pptx', 'slide_index' => 3, 'shape_index' => 1, 'category' => $dto, 'password' => 'password', 'folder' => 'TempSlidesSDK');
        my $chart = $utils->{testSlidesApi}->create_chart_category(%params);
        is(scalar @{$chart->{series}}, 3);
        is(scalar @{$chart->{categories}}, 5);
        is(scalar @{$chart->{series}[0]{data_points}}, 5);
        is($chart->{series}[0]{data_points}[4]{value}, $dto->{data_points}[0]{value});
    };
    if ($@) {
        fail("create_chart_category raised an exception: $@");
    }
    pass();
};

subtest 'chart category update' => sub {
    $utils->initialize('update_chart_category', '');
    eval {
        my $dto = AsposeSlidesCloud::Object::ChartCategory->new();
        $dto->{value} = "New Category";
        my $point1 = AsposeSlidesCloud::Object::OneValueChartDataPoint->new();
        $point1->{value} = 40;
        my $point2 = AsposeSlidesCloud::Object::OneValueChartDataPoint->new();
        $point2->{value} = 50;
        my $point3 = AsposeSlidesCloud::Object::OneValueChartDataPoint->new();
        $point3->{value} = 14;
        my @points = ( $point1, $point2, $point3 );
        $dto->{data_points} = \@points;
        my %params = ('name' => 'test.pptx', 'slide_index' => 3, 'shape_index' => 1, 'category_index' => 2, 'category' => $dto, 'password' => 'password', 'folder' => 'TempSlidesSDK');
        my $chart = $utils->{testSlidesApi}->update_chart_category(%params);
        is(scalar @{$chart->{series}}, 3);
        is(scalar @{$chart->{categories}}, 4);
        is(scalar @{$chart->{series}[0]{data_points}}, 4);
        is($chart->{series}[0]{data_points}[1]{value}, $dto->{data_points}[0]{value});
    };
    if ($@) {
        fail("update_chart_category raised an exception: $@");
    }
    pass();
};

subtest 'chart category delete' => sub {
    $utils->initialize('delete_chart_category', '');
    eval {
        my %params = ('name' => 'test.pptx', 'slide_index' => 3, 'shape_index' => 1, 'category_index' => 2, 'password' => 'password', 'folder' => 'TempSlidesSDK');
        my $chart = $utils->{testSlidesApi}->delete_chart_category(%params);
        is(scalar @{$chart->{series}}, 3);
        is(scalar @{$chart->{categories}}, 3);
        is(scalar @{$chart->{series}[0]{data_points}}, 3);
    };
    if ($@) {
        fail("delete_chart_category raised an exception: $@");
    }
    pass();
};

subtest 'chart data point create' => sub {
    $utils->initialize('create_chart_data_point', '');
    eval {
        my $dto = AsposeSlidesCloud::Object::OneValueChartDataPoint->new();
        $dto->{value} = 40;
        my %params = ('name' => 'test.pptx', 'slide_index' => 3, 'shape_index' => 1, 'series_index' => 2, 'data_point' => $dto, 'password' => 'password', 'folder' => 'TempSlidesSDK');
        $utils->{testSlidesApi}->create_chart_data_point(%params);
    };
    if ($@) {
        if ($@ =~ m/API Exception\((\d+)\): (.*) /s) {
            is($1, 400);
        } else {
            fail("Invalid exception type: $@");
        }
    } else {
        fail("Must have failed because adding data points only works with Scatter & Bubble charts");
    }
    pass();
};

subtest 'chart data point update' => sub {
    $utils->initialize('update_chart_data_point', '');
    eval {
        my $dto = AsposeSlidesCloud::Object::OneValueChartDataPoint->new();
        $dto->{value} = 40;
        my %params = ('name' => 'test.pptx', 'slide_index' => 3, 'shape_index' => 1, 'series_index' => 2, 'point_index' => 2, 'data_point' => $dto, 'password' => 'password', 'folder' => 'TempSlidesSDK');
        my $chart = $utils->{testSlidesApi}->update_chart_data_point(%params);
        is(scalar @{$chart->{series}}, 3);
        is(scalar @{$chart->{categories}}, 4);
        is(scalar @{$chart->{series}[0]{data_points}}, 4);
        is($chart->{series}[1]{data_points}[1]{value}, $dto->{value});
    };
    if ($@) {
        fail("update_chart_data_point raised an exception: $@");
    }
    pass();
};

subtest 'chart data point delete' => sub {
    $utils->initialize('delete_chart_data_point', '');
    eval {
        my %params = ('name' => 'test.pptx', 'slide_index' => 3, 'shape_index' => 1, 'series_index' => 2, 'point_index' => 2, 'password' => 'password', 'folder' => 'TempSlidesSDK');
        my $chart = $utils->{testSlidesApi}->delete_chart_data_point(%params);
        is(scalar @{$chart->{series}}, 3);
        is(scalar @{$chart->{categories}}, 4);
        is($chart->{series}[1]{data_points}[1], undef);
    };
    if ($@) {
        fail("delete_chart_data_point raised an exception: $@");
    }
    pass();
};

subtest 'chart sunburst' => sub {
    $utils->initialize('create_shape', '');
    eval {
        my $dto = AsposeSlidesCloud::Object::Chart->new();
        $dto->{chart_type} = "Sunburst";
        $dto->{width} = 400;
        $dto->{height} = 300;
        my $series1 = AsposeSlidesCloud::Object::OneValueSeries->new();
        $series1->{name} = "Series1";
        my $point1 = AsposeSlidesCloud::Object::OneValueChartDataPoint->new();
        $point1->{value} = 40;
        my $point2 = AsposeSlidesCloud::Object::OneValueChartDataPoint->new();
        $point2->{value} = 50;
        my $point3 = AsposeSlidesCloud::Object::OneValueChartDataPoint->new();
        $point3->{value} = 70;
        my $point4 = AsposeSlidesCloud::Object::OneValueChartDataPoint->new();
        $point4->{value} = 60;
        my @points = ( $point1, $point2, $point3, $point4 );
        $series1->{data_points} = \@points;
        my @series = ( $series1 );
        $dto->{series} = \@series;
        my $category1 = AsposeSlidesCloud::Object::ChartCategory->new();
        $category1->{value} = "Leaf1";
        my @parent_categories1 = ( "Branch1", "Stem1" );
        $category1->{parent_categories} = \@parent_categories1;
        my $category2 = AsposeSlidesCloud::Object::ChartCategory->new();
        $category2->{value} = "Leaf2";
        my @parent_categories2 = ( "Branch1", "Stem1" );
        $category2->{parent_categories} = \@parent_categories2;
        my $category3 = AsposeSlidesCloud::Object::ChartCategory->new();
        $category3->{value} = "Branch2";
        my @parent_categories3 = ( "Stem1" );
        $category3->{parent_categories} = \@parent_categories3;
        my $category4 = AsposeSlidesCloud::Object::ChartCategory->new();
        $category4->{value} = "Stem2";
        my @categories = ( $category1, $category2, $category3, $category4 );
        $dto->{categories} = \@categories;
        my %params = ('name' => 'test.pptx', 'slide_index' => 3, 'dto' => $dto, 'password' => 'password', 'folder' => 'TempSlidesSDK');
        my $chart = $utils->{testSlidesApi}->create_shape(%params);
        is(scalar @{$chart->{series}}, 1);
        is(scalar @{$chart->{categories}}, 4);
    };
    if ($@) {
        fail("create_shape raised an exception: $@");
    }
    pass();
};

subtest 'multi-level category axis' => sub {
    $utils->initialize('create_shape', '');
    eval {
        my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);
        
		my $dto = AsposeSlidesCloud::Object::Chart->new();
        $dto->{chart_type} = "ClusteredColumn";
        $dto->{width} = 500;
        $dto->{height} = 400;
        $dto->{x} = 100;
        $dto->{y} = 100;

        my $series1 = AsposeSlidesCloud::Object::OneValueSeries->new();
        $series1->{type} = "ClusteredColumn";

        my $point1 = AsposeSlidesCloud::Object::OneValueChartDataPoint->new();
        $point1->{value} = 1;
        my $point2 = AsposeSlidesCloud::Object::OneValueChartDataPoint->new();
        $point2->{value} = 2;
        my $point3 = AsposeSlidesCloud::Object::OneValueChartDataPoint->new();
        $point3->{value} = 3;
        my $point4 = AsposeSlidesCloud::Object::OneValueChartDataPoint->new();
        $point4->{value} = 4;
        my $point5 = AsposeSlidesCloud::Object::OneValueChartDataPoint->new();
        $point5->{value} = 5;
        my $point6 = AsposeSlidesCloud::Object::OneValueChartDataPoint->new();
        $point6->{value} = 6;
        my $point7 = AsposeSlidesCloud::Object::OneValueChartDataPoint->new();
        $point7->{value} = 7;
        my $point8 = AsposeSlidesCloud::Object::OneValueChartDataPoint->new();
        $point8->{value} = 8;
        my @points = ( $point1, $point2, $point3, $point4, $point5, $point6, $point7, $point8 );
        $series1->{data_points} = \@points;
        my @series = ( $series1 );
        $dto->{series} = \@series;

        my $category1 = AsposeSlidesCloud::Object::ChartCategory->new();
        $category1->{value} = "Category 1";
        my @parent_categories1 = ( "Sub-category 1", "Root 1" );
        $category1->{parent_categories} = \@parent_categories1;
        my $category2 = AsposeSlidesCloud::Object::ChartCategory->new();
        $category2->{value} = "Category 2";
        my $category3 = AsposeSlidesCloud::Object::ChartCategory->new();
        $category3->{value} = "Category 3";
        my @parent_categories3 = ( "Sub-category 2" );
        $category3->{parent_categories} = \@parent_categories3;
        my $category4 = AsposeSlidesCloud::Object::ChartCategory->new();
        $category4->{value} = "Category 4";
        my $category5 = AsposeSlidesCloud::Object::ChartCategory->new();
        $category5->{value} = "Category 5";
        my @parent_categories5= ( "Sub-category 3", "Root 2" );
        $category5->{parent_categories} = \@parent_categories5;
        my $category6 = AsposeSlidesCloud::Object::ChartCategory->new();
        $category6->{value} = "Category 6";
        my $category7 = AsposeSlidesCloud::Object::ChartCategory->new();
        $category7->{value} = "Category 7";
        my @parent_categories7= ( "Sub-category 4" );
        $category7->{parent_categories} = \@parent_categories7;
        my $category8 = AsposeSlidesCloud::Object::ChartCategory->new();
        $category8->{value} = "Category 8";
        my @categories = ( $category1, $category2, $category3, $category4, $category5, $category6, $category7, $category8 );
        $dto->{categories} = \@categories;

        my %params = ('name' => 'test.pptx', 'slide_index' => 3, 'dto' => $dto, 'password' => 'password', 'folder' => 'TempSlidesSDK');
        my $chart = $utils->{testSlidesApi}->create_shape(%params);
        is(scalar @{$chart->{series}}, 1);
        is(scalar @{$chart->{categories}}, 8);
        is(scalar @{$chart->{categories}[0]{parent_categories}}, 2);
	};
    if ($@) {
        fail("create_shape raised an exception: $@");
    }
    pass();
};

subtest 'hide chart legend' => sub {
    $utils->initialize('get_shape', '');
    $utils->initialize('update_shape', '');
    eval {
        my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);

		my %params = ('name' => 'test.pptx', 'slide_index' => 3, 'shape_index' => 1, 'password' => 'password', 'folder' => 'TempSlidesSDK');
        my $chart = $utils->{testSlidesApi}->get_shape(%params);

        $chart->{legend}->{has_legend} = 0;

        my %params2 = ('name' => 'test.pptx', 'slide_index' => 3, 'shape_index' => 1, 'dto' => $chart, 'password' => 'password', 'folder' => 'TempSlidesSDK');
        $chart = $utils->{testSlidesApi}->update_shape(%params2);
        is($chart->{legend}->{has_legend}, 0);
	};
    if ($@) {
        fail("update_shape raised an exception: $@");
    }
    pass();
};

subtest 'chart grid line format' => sub {
    $utils->initialize('get_shape', '');
    $utils->initialize('update_shape', '');

    eval {
        my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);

		my %params = ('name' => 'test.pptx', 'slide_index' => 3, 'shape_index' => 1, 'password' => 'password', 'folder' => 'TempSlidesSDK');
        my $chart = $utils->{testSlidesApi}->get_shape(%params);

        #horizontal axis
        my $horz_major_line_format_fill_format = AsposeSlidesCloud::Object::NoFill->new();
        my $horz_major_line_format = AsposeSlidesCloud::Object::LineFormat->new();
        $horz_major_line_format->{fill_format} = $horz_major_line_format_fill_format;
        my $horz_major_line = AsposeSlidesCloud::Object::ChartLinesFormat->new();
        $horz_major_line->{line_format} = $horz_major_line_format;

        my $horz_minor_line_format_fill_format = AsposeSlidesCloud::Object::SolidFill->new();
        $horz_minor_line_format_fill_format->{color} = 'Black';
        my $horz_minor_line_format = AsposeSlidesCloud::Object::LineFormat->new();
        $horz_minor_line_format->{fill_format} = $horz_minor_line_format_fill_format;
        my $horz_minor_line = AsposeSlidesCloud::Object::ChartLinesFormat->new();
        $horz_minor_line->{line_format} = $horz_minor_line_format;

        my $horz_axis = AsposeSlidesCloud::Object::Axis->new();
        $horz_axis->{major_grid_lines_format} = $horz_major_line;
        $horz_axis->{minor_grid_lines_format} = $horz_minor_line;

        #vertical axis
        my $gradient_stop1 = AsposeSlidesCloud::Object::GradientFillStop->new();
        $gradient_stop1->{color} = 'White';
        $gradient_stop1->{position} = 0;
        my $gradient_stop2 = AsposeSlidesCloud::Object::GradientFillStop->new();
        $gradient_stop2->{color} = 'Black';
        $gradient_stop2->{position} = 1;
        my @gradient_stops = ($gradient_stop1, $gradient_stop2);

        my $vert_major_line_format_fill_format = AsposeSlidesCloud::Object::GradientFill->new();
        $vert_major_line_format_fill_format->{direction} = 'FromCorner1';
        $vert_major_line_format_fill_format->{stops} = \@gradient_stops;
        my $vert_major_line_format = AsposeSlidesCloud::Object::LineFormat->new();
        $vert_major_line_format->{fill_format} = $vert_major_line_format_fill_format;
        my $vert_major_line = AsposeSlidesCloud::Object::ChartLinesFormat->new();
        $vert_major_line->{line_format} = $vert_major_line_format;

        my $vert_minor_line_format_fill_format = AsposeSlidesCloud::Object::NoFill->new();
        my $vert_minor_line_format = AsposeSlidesCloud::Object::LineFormat->new();
        $vert_minor_line_format->{fill_format} = $vert_minor_line_format_fill_format;
        my $vert_minor_line = AsposeSlidesCloud::Object::ChartLinesFormat->new();
        $vert_minor_line->{line_format} = $vert_minor_line_format;

        my $vert_axis = AsposeSlidesCloud::Object::Axis->new();
        $vert_axis->{major_grid_lines_format} = $vert_major_line;
        $vert_axis->{minor_grid_lines_format} = $vert_minor_line;

        my $axes = AsposeSlidesCloud::Object::Axes->new();
        $axes->{horizontal_axis} = $horz_axis;
        $axes->{vertical_axis} = $vert_axis;

        $chart->{axes} = $axes;

        my %params2 = ('name' => 'test.pptx', 'slide_index' => 3, 'shape_index' => 1, 'dto' => $chart, 'password' => 'password', 'folder' => 'TempSlidesSDK');
        $chart = $utils->{testSlidesApi}->update_shape(%params2);
        
        is($chart->{axes}->{horizontal_axis}->{major_grid_lines_format}->{line_format}->{fill_format}->{type}, 'NoFill');
	    is($chart->{axes}->{horizontal_axis}->{minor_grid_lines_format}->{line_format}->{fill_format}->{type}, 'Solid');
	    is($chart->{axes}->{vertical_axis}->{major_grid_lines_format}->{line_format}->{fill_format}->{type}, 'Gradient');
	    is($chart->{axes}->{vertical_axis}->{minor_grid_lines_format}->{line_format}->{fill_format}->{type}, 'NoFill');
    };
    if ($@) {
        fail("update_shape raised an exception: $@");
    }
    pass();
};

subtest 'chart series groups' => sub {
    $utils->initialize('get_shape', '');
    $utils->initialize('set_chart_series_group', '');
    eval {
        my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);

		my %params = ('name' => 'test.pptx', 'slide_index' => 3, 'shape_index' => 1, 'password' => 'password', 'folder' => 'TempSlidesSDK');
        my $chart = $utils->{testSlidesApi}->get_shape(%params);

        my $series_group = $chart->{series_groups}[0];
        $series_group->{overlap} = 10;
        
        my %params2 = ('name' => 'test.pptx', 'slide_index' => 3, 'shape_index' => 1, 'series_group_index' => 1, 'series_group' => $series_group, 'password' => 'password', 'folder' => 'TempSlidesSDK');
        $chart = $utils->{testSlidesApi}->set_chart_series_group(%params2);
        is ($chart->{series_groups}[0]->{overlap}, 10);
	};
    if ($@) {
        fail("set_chart_series_group raised an exception: $@");
    }
    pass();
};

subtest 'set chart legend' => sub {
    $utils->initialize('set_chart_legend', '');
    eval {
        my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);

        my $legend = AsposeSlidesCloud::Object::Legend->new();
        $legend->{overlay} = 1;
        my $fill_format = AsposeSlidesCloud::Object::SolidFill->new();
        $fill_format->{color} = "#77CEF9";
        $legend->{fill_format} = $fill_format;

        my %params = ('name' => 'test.pptx', 'slide_index' => 3, 'shape_index' => 1, 'legend' => $legend, 'password' => 'password', 'folder' => 'TempSlidesSDK');
        my $response = $utils->{testSlidesApi}->set_chart_legend(%params);
        is ($legend->{overlay}, 1);
        is ($legend->{fill_format}->{type}, 'Solid');
	};
    if ($@) {
        fail("set_chart_legend raised an exception: $@");
    }
    pass();
};

subtest 'set chart axis' => sub {
    $utils->initialize('set_chart_axis', '');
    eval {
        my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);

        my $axis = AsposeSlidesCloud::Object::Axis->new();
        $axis->{has_title} = 1;
        $axis->{is_automatic_max_value} = 0;
        $axis->{max_value} = 10;

        my %params = ('name' => 'test.pptx', 'slide_index' => 3, 'shape_index' => 1, "axis_type" => 'VerticalAxis', 'axis'=>$axis, 'password' => 'password', 'folder' => 'TempSlidesSDK');
        my $response = $utils->{testSlidesApi}->set_chart_axis(%params);
        is ($response->{has_title}, $axis->{has_title});
        is ($response->{is_automatic_max_value}, $axis->{is_automatic_max_value});
        is ($response->{max_value}, $axis->{max_value});
	};
    if ($@) {
        fail("set_chart_axis raised an exception: $@");
    }
    pass();
};

subtest 'set chart wall' => sub {
    $utils->initialize('set_chart_wall', '');
    eval {
        my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);

        my $chart_wall = AsposeSlidesCloud::Object::ChartWall->new();
        my $fill_format = AsposeSlidesCloud::Object::SolidFill->new();
        $fill_format->{color} = "#77CEF9";
        $chart_wall->{fill_format} = $fill_format;

        my %params = ('name' => 'test.pptx', 'slide_index' => 8, 'shape_index' => 2, "chart_wall_type" => 'BackWall', 'chart_wall'=>$chart_wall, 'password' => 'password', 'folder' => 'TempSlidesSDK');
        my $response = $utils->{testSlidesApi}->set_chart_wall(%params);
        is ($response->{fill_format}->{type}, 'Solid');
	};
    if ($@) {
        fail("set_chart_wall raised an exception: $@");
    }
    pass();
};

subtest 'update data point format' => sub {
    $utils->initialize('update_chart_data_point', '');
    eval {
        my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);

        my $fill_format = AsposeSlidesCloud::Object::SolidFill->new();
        $fill_format->{color} = '#77CEF9';
        my $line_format = AsposeSlidesCloud::Object::LineFormat->new();
        my $line_format_fill_format = AsposeSlidesCloud::Object::SolidFill->new();
        $line_format_fill_format->{color} = "#E85052";
        $line_format->{fill_format} = $line_format_fill_format;
        my $effect_format = AsposeSlidesCloud::Object::EffectFormat->new();
        my $blur = AsposeSlidesCloud::Object::BlurEffect->new();
        $blur->{grow} = 1;
        $blur->{radius} = 5;
        $effect_format->{blur} = $blur;

        my $dto = AsposeSlidesCloud::Object::OneValueChartDataPoint->new();
        $dto->{value} = 40;
        $dto->{fill_format} = $fill_format;
        $dto->{line_format} = $line_format;
        $dto->{effect_format} = $effect_format;

        my %params = ('name' => 'test.pptx', 'slide_index' => 8, 'shape_index' => 2, "series_index" => 2, 'point_index' => 2, 'data_point' => $dto, 'password' => 'password', 'folder' => 'TempSlidesSDK');
        my $chart = $utils->{testSlidesApi}->update_chart_data_point(%params);
        my $data_point = $chart->{series}[1]{data_points}[1];
        is ($data_point->{fill_format}->{type}, 'Solid');
        is ($data_point->{line_format}->{fill_format}->{type}, 'Solid');
        
	};
    if ($@) {
        fail("update_chart_data_point raised an exception: $@");
    }
    pass();
};

subtest 'chart formulas' => sub {
    $utils->initialize('get_shape', '');
    eval {
        my $dto = AsposeSlidesCloud::Object::Chart->new();
        $dto->{chart_type} = "ClusteredColumn";
        $dto->{width} = 400;
        $dto->{height} = 300;
        my $data_source_for_categories = AsposeSlidesCloud::Object::Workbook->new();
        $data_source_for_categories->{worksheet_index} = 1;
        $data_source_for_categories->{column_index} = 1;
        $data_source_for_categories->{row_index} = 2;
        $dto->{data_source_for_categories} = $data_source_for_categories;
        my $category1 = AsposeSlidesCloud::Object::ChartCategory->new();
        $category1->{value} = "Category1";
        my $category2 = AsposeSlidesCloud::Object::ChartCategory->new();
        $category2->{value} = "Category2";
        my $category3 = AsposeSlidesCloud::Object::ChartCategory->new();
        $category3->{value} = "Category3";
        my @categories = ( $category1, $category2, $category3 );
        $dto->{categories} = \@categories;

        my $series1 = AsposeSlidesCloud::Object::OneValueSeries->new();
        my $data_source_for_series1_name = AsposeSlidesCloud::Object::Workbook->new();
        $data_source_for_series1_name->{worksheet_index} = 1;
        $data_source_for_series1_name->{column_index} = 2;
        $data_source_for_series1_name->{row_index} = 1;
        $series1->{data_source_for_series_name} = $data_source_for_series1_name;
        $series1->{name} = "Series1";
        my $data_source_for_series1_values = AsposeSlidesCloud::Object::Workbook->new();
        $data_source_for_series1_values->{worksheet_index} = 1;
        $data_source_for_series1_values->{column_index} = 2;
        $data_source_for_series1_values->{row_index} = 2;
        $series1->{data_source_for_values} = $data_source_for_series1_values;
        my $point11 = AsposeSlidesCloud::Object::OneValueChartDataPoint->new();
        $point11->{value} = 40;
        my $point12 = AsposeSlidesCloud::Object::OneValueChartDataPoint->new();
        $point12->{value} = 50;
        my $point13 = AsposeSlidesCloud::Object::OneValueChartDataPoint->new();
        $point13->{value_formula} = "SUM(B2:B3)";
        my @points1 = ( $point11, $point12, $point13 );
        $series1->{data_points} = \@points1;
        my @series = ( $series1 );
        $dto->{series} = \@series;
        
        my %params = ('name' => 'test.pptx', 'slide_index' => 3, 'dto' => $dto, 'password' => 'password', 'folder' => 'TempSlidesSDK');
        my $chart = $utils->{testSlidesApi}->create_shape(%params);
        is($chart->{series}[0]{data_points}[2]->{value}, 90);
    };
    if ($@) {
        fail("get_shape raised an exception: $@");
    }
    pass();
};

subtest 'import chart from workbook' => sub {
    eval {
        my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);

        my $source = read_file("TestData/oleObject.xlsx", { binmode => ':raw' });
        my %params = (
            'name' => "test.pptx",
            'slide_index' => 3,
            'worksheet_name' => "Sheet1",
            'document' => $source,
            'chart_index' => 1,
            'password' => "password",
            'folder' => "TempSlidesSDK");
        my $result = $utils->{testSlidesApi}->import_chart_from_workbook(%params);
        is($result->{type}, "Chart");
    };
    if ($@) {
        fail("import_chart_from_workbook raised an exception: $@");
    }
    pass();
};

subtest 'import chart from workbook by path' => sub {
    eval {
        my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);

        my $source = read_file("TestData/oleObject.xlsx", { binmode => ':raw' });
        my %upload_params = ('path' => "TempSlidesSDK/oleObject.xlsx", 'file' => $source);
        $utils->{testSlidesApi}->upload_file(%upload_params);

        my %params = (
            'name' => "test.pptx",
            'slide_index' => 3,
            'worksheet_name' => "Sheet1",
            'chart_index' => 1,
            'workbook_path' => "TempSlidesSDK/oleObject.xlsx",
            'password' => "password",
            'folder' => "TempSlidesSDK");
        my $result = $utils->{testSlidesApi}->import_chart_from_workbook(%params);
        is($result->{type}, "Chart");
    };
    if ($@) {
        fail("import_chart_from_workbook raised an exception: $@");
    }
    pass();
};

done_testing;
