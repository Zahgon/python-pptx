"""Composers for default chart XML for various chart types."""

from __future__ import annotations

from copy import deepcopy
from xml.sax.saxutils import escape

from pptx.enum.chart import XL_CHART_TYPE
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls


def ChartXmlWriter(chart_type, chart_data):
    """
    Factory function returning appropriate XML writer object for
    *chart_type*, loaded with *chart_type* and *chart_data*.
    """
    XL_CT = XL_CHART_TYPE
    try:
        BuilderCls = {
            XL_CT.AREA: _AreaChartXmlWriter,
            XL_CT.AREA_STACKED: _AreaChartXmlWriter,
            XL_CT.AREA_STACKED_100: _AreaChartXmlWriter,
            XL_CT.BAR_CLUSTERED: _BarChartXmlWriter,
            XL_CT.BAR_STACKED: _BarChartXmlWriter,
            XL_CT.BAR_STACKED_100: _BarChartXmlWriter,
            XL_CT.BUBBLE: _BubbleChartXmlWriter,
            XL_CT.BUBBLE_THREE_D_EFFECT: _BubbleChartXmlWriter,
            XL_CT.COLUMN_CLUSTERED: _BarChartXmlWriter,
            XL_CT.COLUMN_STACKED: _BarChartXmlWriter,
            XL_CT.COLUMN_STACKED_100: _BarChartXmlWriter,
            XL_CT.DOUGHNUT: _DoughnutChartXmlWriter,
            XL_CT.DOUGHNUT_EXPLODED: _DoughnutChartXmlWriter,
            XL_CT.LINE: _LineChartXmlWriter,
            XL_CT.LINE_MARKERS: _LineChartXmlWriter,
            XL_CT.LINE_MARKERS_STACKED: _LineChartXmlWriter,
            XL_CT.LINE_MARKERS_STACKED_100: _LineChartXmlWriter,
            XL_CT.LINE_STACKED: _LineChartXmlWriter,
            XL_CT.LINE_STACKED_100: _LineChartXmlWriter,
            XL_CT.PIE: _PieChartXmlWriter,
            XL_CT.PIE_EXPLODED: _PieChartXmlWriter,
            XL_CT.RADAR: _RadarChartXmlWriter,
            XL_CT.RADAR_FILLED: _RadarChartXmlWriter,
            XL_CT.RADAR_MARKERS: _RadarChartXmlWriter,
            XL_CT.XY_SCATTER: _XyChartXmlWriter,
            XL_CT.XY_SCATTER_LINES: _XyChartXmlWriter,
            XL_CT.XY_SCATTER_LINES_NO_MARKERS: _XyChartXmlWriter,
            XL_CT.XY_SCATTER_SMOOTH: _XyChartXmlWriter,
            XL_CT.XY_SCATTER_SMOOTH_NO_MARKERS: _XyChartXmlWriter,
        }[chart_type]
    except KeyError:
        raise NotImplementedError("XML writer for chart type %s not yet implemented" % chart_type)
    return BuilderCls(chart_type, chart_data)


def SeriesXmlRewriterFactory(chart_type, chart_data):
    """
    Return a |_BaseSeriesXmlRewriter| subclass appropriate to *chart_type*.
    """
    pass


class _BaseChartXmlWriter(object):
    """
    Generates XML text (unicode) for a default chart, like the one added by
    PowerPoint when you click the *Add Column Chart* button on the ribbon.
    Differentiated XML for different chart types is provided by subclasses.
    """

    def __init__(self, chart_type, series_seq):
        super(_BaseChartXmlWriter, self).__init__()
        self._chart_type = chart_type
        self._chart_data = series_seq
        self._series_seq = list(series_seq)

    @property
    def xml(self):
        """
        The full XML stream for the chart specified by this chart builder, as
        unicode text. This method must be overridden by each subclass.
        """
        raise NotImplementedError("must be implemented by all subclasses")


class _BaseSeriesXmlWriter(object):
    """
    Provides shared members for series XML writers.
    """

    def __init__(self, series, date_1904=False):
        super(_BaseSeriesXmlWriter, self).__init__()
        self._series = series
        self._date_1904 = date_1904

    @property
    def name(self):
        """
        The XML-escaped name for this series.
        """
        pass

    def numRef_xml(self, wksht_ref, number_format, values):
        """
        Return the ``<c:numRef>`` element specified by the parameters as
        unicode text.
        """
        pass

    def pt_xml(self, values):
        """
        Return the ``<c:ptCount>`` and sequence of ``<c:pt>`` elements
        corresponding to *values* as a single unicode text string.
        `c:ptCount` refers to the number of `c:pt` elements in this sequence.
        The `idx` attribute value for `c:pt` elements locates the data point
        in the overall data point sequence of the chart and is started at
        *offset*.
        """
        pass

    @property
    def tx(self):
        """
        Return a ``<c:tx>`` oxml element for this series, containing the
        series name.
        """
        pass

    @property
    def tx_xml(self):
        """
        Return the ``<c:tx>`` (tx is short for 'text') element for this
        series as unicode text. This element contains the series name.
        """
        pass

    @property
    def _tx_tmpl(self):
        """
        The string formatting template for the ``<c:tx>`` element for this
        series, containing the series title and spreadsheet range reference.
        """
        pass


class _BaseSeriesXmlRewriter(object):
    """
    Base class for series XML rewriters.
    """

    def __init__(self, chart_data):
        super(_BaseSeriesXmlRewriter, self).__init__()
        self._chart_data = chart_data

    def replace_series_data(self, chartSpace):
        """
        Rewrite the series data under *chartSpace* using the chart data
        contents. All series-level formatting is left undisturbed. If
        the chart data contains fewer series than *chartSpace*, the extra
        series in *chartSpace* are deleted. If *chart_data* contains more
        series than the *chartSpace* element, new series are added to the
        last plot in the chart and series formatting is "cloned" from the
        last series in that plot.
        """
        pass

    def _add_cloned_sers(self, plotArea, count):
        """
        Add `c:ser` elements to the last xChart element in *plotArea*, cloned
        from the last `c:ser` child of that last xChart.
        """
        pass

    def _adjust_ser_count(self, plotArea, new_ser_count):
        """
        Adjust the number of c:ser elements in *plotArea* to *new_ser_count*.
        Excess c:ser elements are deleted from the end, along with any xChart
        elements that are left empty as a result. Series elements are
        considered in xChart + series order. Any new c:ser elements required
        are added to the last xChart element and cloned from the last c:ser
        element in that xChart.
        """
        pass

    def _rewrite_ser_data(self, ser, series_data, date_1904):
        """
        Rewrite selected child elements of *ser* based on the values in
        *series_data*.
        """
        raise NotImplementedError("must be implemented by each subclass")

    def _trim_ser_count_by(self, plotArea, count):
        """
        Remove the last *count* ser elements from *plotArea*. Any xChart
        elements having no ser child elements after trimming are also
        removed.
        """
        pass


class _AreaChartXmlWriter(_BaseChartXmlWriter):
    """
    Provides specialized methods particular to the ``<c:areaChart>`` element.
    """

    @property
    def xml(self):
        return (
            "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>\n"
            '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawin'
            'gml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/draw'
            'ingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/off'
            'iceDocument/2006/relationships">\n'
            '  <c:date1904 val="0"/>\n'
            '  <c:roundedCorners val="0"/>\n'
            "  <c:chart>\n"
            '    <c:autoTitleDeleted val="0"/>\n'
            "    <c:plotArea>\n"
            "      <c:layout/>\n"
            "      <c:areaChart>\n"
            "{grouping_xml}"
            '        <c:varyColors val="0"/>\n'
            "{ser_xml}"
            "        <c:dLbls>\n"
            '          <c:showLegendKey val="0"/>\n'
            '          <c:showVal val="0"/>\n'
            '          <c:showCatName val="0"/>\n'
            '          <c:showSerName val="0"/>\n'
            '          <c:showPercent val="0"/>\n'
            '          <c:showBubbleSize val="0"/>\n'
            "        </c:dLbls>\n"
            '        <c:axId val="-2101159928"/>\n'
            '        <c:axId val="-2100718248"/>\n'
            "      </c:areaChart>\n"
            "{cat_ax_xml}"
            "      <c:valAx>\n"
            '        <c:axId val="-2100718248"/>\n'
            "        <c:scaling>\n"
            '          <c:orientation val="minMax"/>\n'
            "        </c:scaling>\n"
            '        <c:delete val="0"/>\n'
            '        <c:axPos val="l"/>\n'
            "        <c:majorGridlines/>\n"
            '        <c:numFmt formatCode="General" sourceLinked="1"/>\n'
            '        <c:majorTickMark val="out"/>\n'
            '        <c:minorTickMark val="none"/>\n'
            '        <c:tickLblPos val="nextTo"/>\n'
            '        <c:crossAx val="-2101159928"/>\n'
            '        <c:crosses val="autoZero"/>\n'
            '        <c:crossBetween val="midCat"/>\n'
            "      </c:valAx>\n"
            "    </c:plotArea>\n"
            "    <c:legend>\n"
            '      <c:legendPos val="r"/>\n'
            "      <c:layout/>\n"
            '      <c:overlay val="0"/>\n'
            "    </c:legend>\n"
            '    <c:plotVisOnly val="1"/>\n'
            '    <c:dispBlanksAs val="zero"/>\n'
            '    <c:showDLblsOverMax val="0"/>\n'
            "  </c:chart>\n"
            "  <c:txPr>\n"
            "    <a:bodyPr/>\n"
            "    <a:lstStyle/>\n"
            "    <a:p>\n"
            "      <a:pPr>\n"
            '        <a:defRPr sz="1800"/>\n'
            "      </a:pPr>\n"
            "      <a:endParaRPr/>\n"
            "    </a:p>\n"
            "  </c:txPr>\n"
            "</c:chartSpace>\n"
        ).format(
            **{
                "grouping_xml": self._grouping_xml,
                "ser_xml": self._ser_xml,
                "cat_ax_xml": self._cat_ax_xml,
            }
        )

    @property
    def _cat_ax_xml(self):
        pass

    @property
    def _grouping_xml(self):
        pass

    @property
    def _ser_xml(self):
        pass


class _BarChartXmlWriter(_BaseChartXmlWriter):
    """
    Provides specialized methods particular to the ``<c:barChart>`` element.
    """

    @property
    def xml(self):
        return (
            "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>\n"
            '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawin'
            'gml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/draw'
            'ingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/off'
            'iceDocument/2006/relationships">\n'
            '  <c:date1904 val="0"/>\n'
            "  <c:chart>\n"
            '    <c:autoTitleDeleted val="0"/>\n'
            "    <c:plotArea>\n"
            "      <c:barChart>\n"
            "{barDir_xml}"
            "{grouping_xml}"
            "{ser_xml}"
            "{overlap_xml}"
            '        <c:axId val="-2068027336"/>\n'
            '        <c:axId val="-2113994440"/>\n'
            "      </c:barChart>\n"
            "{cat_ax_xml}"
            "      <c:valAx>\n"
            '        <c:axId val="-2113994440"/>\n'
            "        <c:scaling/>\n"
            '        <c:delete val="0"/>\n'
            '        <c:axPos val="{val_ax_pos}"/>\n'
            "        <c:majorGridlines/>\n"
            '        <c:majorTickMark val="out"/>\n'
            '        <c:minorTickMark val="none"/>\n'
            '        <c:tickLblPos val="nextTo"/>\n'
            '        <c:crossAx val="-2068027336"/>\n'
            '        <c:crosses val="autoZero"/>\n'
            "      </c:valAx>\n"
            "    </c:plotArea>\n"
            '    <c:dispBlanksAs val="gap"/>\n'
            "  </c:chart>\n"
            "  <c:txPr>\n"
            "    <a:bodyPr/>\n"
            "    <a:lstStyle/>\n"
            "    <a:p>\n"
            "      <a:pPr>\n"
            '        <a:defRPr sz="1800"/>\n'
            "      </a:pPr>\n"
            '      <a:endParaRPr lang="en-US"/>\n'
            "    </a:p>\n"
            "  </c:txPr>\n"
            "</c:chartSpace>\n"
        ).format(
            **{
                "barDir_xml": self._barDir_xml,
                "grouping_xml": self._grouping_xml,
                "ser_xml": self._ser_xml,
                "overlap_xml": self._overlap_xml,
                "cat_ax_xml": self._cat_ax_xml,
                "val_ax_pos": self._val_ax_pos,
            }
        )

    @property
    def _barDir_xml(self):
        pass

    @property
    def _cat_ax_pos(self):
        pass

    @property
    def _cat_ax_xml(self):
        pass

    @property
    def _grouping_xml(self):
        pass

    @property
    def _overlap_xml(self):
        pass

    @property
    def _ser_xml(self):
        pass

    @property
    def _val_ax_pos(self):
        pass


class _DoughnutChartXmlWriter(_BaseChartXmlWriter):
    """
    Provides specialized methods particular to the ``<c:doughnutChart>``
    element.
    """

    @property
    def xml(self):
        return (
            "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>\n"
            '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawin'
            'gml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/draw'
            'ingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/off'
            'iceDocument/2006/relationships">\n'
            '  <c:date1904 val="0"/>\n'
            '  <c:roundedCorners val="0"/>\n'
            "  <c:chart>\n"
            '    <c:autoTitleDeleted val="0"/>\n'
            "    <c:plotArea>\n"
            "      <c:layout/>\n"
            "      <c:doughnutChart>\n"
            '        <c:varyColors val="1"/>\n'
            "{ser_xml}"
            "        <c:dLbls>\n"
            '          <c:showLegendKey val="0"/>\n'
            '          <c:showVal val="0"/>\n'
            '          <c:showCatName val="0"/>\n'
            '          <c:showSerName val="0"/>\n'
            '          <c:showPercent val="0"/>\n'
            '          <c:showBubbleSize val="0"/>\n'
            '          <c:showLeaderLines val="1"/>\n'
            "        </c:dLbls>\n"
            '        <c:firstSliceAng val="0"/>\n'
            '        <c:holeSize val="50"/>\n'
            "      </c:doughnutChart>\n"
            "    </c:plotArea>\n"
            "    <c:legend>\n"
            '      <c:legendPos val="r"/>\n'
            "      <c:layout/>\n"
            '      <c:overlay val="0"/>\n'
            "    </c:legend>\n"
            '    <c:plotVisOnly val="1"/>\n'
            '    <c:dispBlanksAs val="gap"/>\n'
            '    <c:showDLblsOverMax val="0"/>\n'
            "  </c:chart>\n"
            "  <c:txPr>\n"
            "    <a:bodyPr/>\n"
            "    <a:lstStyle/>\n"
            "    <a:p>\n"
            "      <a:pPr>\n"
            '        <a:defRPr sz="1800"/>\n'
            "      </a:pPr>\n"
            "      <a:endParaRPr/>\n"
            "    </a:p>\n"
            "  </c:txPr>\n"
            "</c:chartSpace>\n"
        ).format(**{"ser_xml": self._ser_xml})

    @property
    def _explosion_xml(self):
        pass

    @property
    def _ser_xml(self):
        pass


class _LineChartXmlWriter(_BaseChartXmlWriter):
    """
    Provides specialized methods particular to the ``<c:lineChart>`` element.
    """

    @property
    def xml(self):
        return (
            "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>\n"
            '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawin'
            'gml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/draw'
            'ingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/off'
            'iceDocument/2006/relationships">\n'
            '  <c:date1904 val="0"/>\n'
            "  <c:chart>\n"
            '    <c:autoTitleDeleted val="0"/>\n'
            "    <c:plotArea>\n"
            "      <c:lineChart>\n"
            "{grouping_xml}"
            '        <c:varyColors val="0"/>\n'
            "{ser_xml}"
            '        <c:marker val="1"/>\n'
            '        <c:smooth val="0"/>\n'
            '        <c:axId val="2118791784"/>\n'
            '        <c:axId val="2140495176"/>\n'
            "      </c:lineChart>\n"
            "{cat_ax_xml}"
            "      <c:valAx>\n"
            '        <c:axId val="2140495176"/>\n'
            "        <c:scaling/>\n"
            '        <c:delete val="0"/>\n'
            '        <c:axPos val="l"/>\n'
            "        <c:majorGridlines/>\n"
            '        <c:majorTickMark val="out"/>\n'
            '        <c:minorTickMark val="none"/>\n'
            '        <c:tickLblPos val="nextTo"/>\n'
            '        <c:crossAx val="2118791784"/>\n'
            '        <c:crosses val="autoZero"/>\n'
            "      </c:valAx>\n"
            "    </c:plotArea>\n"
            "    <c:legend>\n"
            '      <c:legendPos val="r"/>\n'
            "      <c:layout/>\n"
            '      <c:overlay val="0"/>\n'
            "    </c:legend>\n"
            '    <c:plotVisOnly val="1"/>\n'
            '    <c:dispBlanksAs val="gap"/>\n'
            '    <c:showDLblsOverMax val="0"/>\n'
            "  </c:chart>\n"
            "  <c:txPr>\n"
            "    <a:bodyPr/>\n"
            "    <a:lstStyle/>\n"
            "    <a:p>\n"
            "      <a:pPr>\n"
            '        <a:defRPr sz="1800"/>\n'
            "      </a:pPr>\n"
            '      <a:endParaRPr lang="en-US"/>\n'
            "    </a:p>\n"
            "  </c:txPr>\n"
            "</c:chartSpace>\n"
        ).format(
            **{
                "grouping_xml": self._grouping_xml,
                "ser_xml": self._ser_xml,
                "cat_ax_xml": self._cat_ax_xml,
            }
        )

    @property
    def _cat_ax_xml(self):
        pass

    @property
    def _grouping_xml(self):
        pass

    @property
    def _marker_xml(self):
        pass

    @property
    def _ser_xml(self):
        pass


class _PieChartXmlWriter(_BaseChartXmlWriter):
    """
    Provides specialized methods particular to the ``<c:pieChart>`` element.
    """

    @property
    def xml(self):
        return (
            "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>\n"
            '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawin'
            'gml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/draw'
            'ingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/off'
            'iceDocument/2006/relationships">\n'
            "  <c:chart>\n"
            '    <c:autoTitleDeleted val="0"/>\n'
            "    <c:plotArea>\n"
            "      <c:pieChart>\n"
            '        <c:varyColors val="1"/>\n'
            "{ser_xml}"
            "      </c:pieChart>\n"
            "    </c:plotArea>\n"
            '    <c:dispBlanksAs val="gap"/>\n'
            "  </c:chart>\n"
            "  <c:txPr>\n"
            "    <a:bodyPr/>\n"
            "    <a:lstStyle/>\n"
            "    <a:p>\n"
            "      <a:pPr>\n"
            '        <a:defRPr sz="1800"/>\n'
            "      </a:pPr>\n"
            '      <a:endParaRPr lang="en-US"/>\n'
            "    </a:p>\n"
            "  </c:txPr>\n"
            "</c:chartSpace>\n"
        ).format(**{"ser_xml": self._ser_xml})

    @property
    def _explosion_xml(self):
        pass

    @property
    def _ser_xml(self):
        pass


class _RadarChartXmlWriter(_BaseChartXmlWriter):
    """
    Generates XML for the ``<c:radarChart>`` element.
    """

    @property
    def xml(self):
        return (
            "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>\n"
            '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawin'
            'gml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/draw'
            'ingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/off'
            'iceDocument/2006/relationships">\n'
            '  <c:date1904 val="0"/>\n'
            '  <c:roundedCorners val="0"/>\n'
            '  <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.'
            'org/markup-compatibility/2006">\n'
            '    <mc:Choice xmlns:c14="http://schemas.microsoft.com/office/d'
            'rawing/2007/8/2/chart" Requires="c14">\n'
            '      <c14:style val="118"/>\n'
            "    </mc:Choice>\n"
            "    <mc:Fallback>\n"
            '      <c:style val="18"/>\n'
            "    </mc:Fallback>\n"
            "  </mc:AlternateContent>\n"
            "  <c:chart>\n"
            '    <c:autoTitleDeleted val="0"/>\n'
            "    <c:plotArea>\n"
            "      <c:layout/>\n"
            "      <c:radarChart>\n"
            '        <c:radarStyle val="{radar_style}"/>\n'
            '        <c:varyColors val="0"/>\n'
            "{ser_xml}"
            '        <c:axId val="2073612648"/>\n'
            '        <c:axId val="-2112772216"/>\n'
            "      </c:radarChart>\n"
            "      <c:catAx>\n"
            '        <c:axId val="2073612648"/>\n'
            "        <c:scaling>\n"
            '          <c:orientation val="minMax"/>\n'
            "        </c:scaling>\n"
            '        <c:delete val="0"/>\n'
            '        <c:axPos val="b"/>\n'
            "        <c:majorGridlines/>\n"
            '        <c:numFmt formatCode="m/d/yy" sourceLinked="1"/>\n'
            '        <c:majorTickMark val="out"/>\n'
            '        <c:minorTickMark val="none"/>\n'
            '        <c:tickLblPos val="nextTo"/>\n'
            '        <c:crossAx val="-2112772216"/>\n'
            '        <c:crosses val="autoZero"/>\n'
            '        <c:auto val="1"/>\n'
            '        <c:lblAlgn val="ctr"/>\n'
            '        <c:lblOffset val="100"/>\n'
            '        <c:noMultiLvlLbl val="0"/>\n'
            "      </c:catAx>\n"
            "      <c:valAx>\n"
            '        <c:axId val="-2112772216"/>\n'
            "        <c:scaling>\n"
            '          <c:orientation val="minMax"/>\n'
            "        </c:scaling>\n"
            '        <c:delete val="0"/>\n'
            '        <c:axPos val="l"/>\n'
            "        <c:majorGridlines/>\n"
            '        <c:numFmt formatCode="General" sourceLinked="1"/>\n'
            '        <c:majorTickMark val="cross"/>\n'
            '        <c:minorTickMark val="none"/>\n'
            '        <c:tickLblPos val="nextTo"/>\n'
            '        <c:crossAx val="2073612648"/>\n'
            '        <c:crosses val="autoZero"/>\n'
            '        <c:crossBetween val="between"/>\n'
            "      </c:valAx>\n"
            "    </c:plotArea>\n"
            '    <c:plotVisOnly val="1"/>\n'
            '    <c:dispBlanksAs val="gap"/>\n'
            '    <c:showDLblsOverMax val="0"/>\n'
            "  </c:chart>\n"
            "  <c:txPr>\n"
            "    <a:bodyPr/>\n"
            "    <a:lstStyle/>\n"
            "    <a:p>\n"
            "      <a:pPr>\n"
            '        <a:defRPr sz="1800"/>\n'
            "      </a:pPr>\n"
            '      <a:endParaRPr lang="en-US"/>\n'
            "    </a:p>\n"
            "  </c:txPr>\n"
            "</c:chartSpace>\n"
        ).format(**{"radar_style": self._radar_style, "ser_xml": self._ser_xml})

    @property
    def _marker_xml(self):
        pass

    @property
    def _radar_style(self):
        pass

    @property
    def _ser_xml(self):
        pass


class _XyChartXmlWriter(_BaseChartXmlWriter):
    """
    Generates XML for the ``<c:scatterChart>`` element.
    """

    @property
    def xml(self):
        xml = (
            "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>\n"
            '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawin'
            'gml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/draw'
            'ingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/off'
            'iceDocument/2006/relationships">\n'
            "  <c:chart>\n"
            "    <c:plotArea>\n"
            "      <c:scatterChart>\n"
            '        <c:scatterStyle val="%s"/>\n'
            '        <c:varyColors val="0"/>\n'
            "%s"
            '        <c:axId val="-2128940872"/>\n'
            '        <c:axId val="-2129643912"/>\n'
            "      </c:scatterChart>\n"
            "      <c:valAx>\n"
            '        <c:axId val="-2128940872"/>\n'
            "        <c:scaling>\n"
            '          <c:orientation val="minMax"/>\n'
            "        </c:scaling>\n"
            '        <c:delete val="0"/>\n'
            '        <c:axPos val="b"/>\n'
            '        <c:numFmt formatCode="General" sourceLinked="1"/>\n'
            '        <c:majorTickMark val="out"/>\n'
            '        <c:minorTickMark val="none"/>\n'
            '        <c:tickLblPos val="nextTo"/>\n'
            '        <c:crossAx val="-2129643912"/>\n'
            '        <c:crosses val="autoZero"/>\n'
            '        <c:crossBetween val="midCat"/>\n'
            "      </c:valAx>\n"
            "      <c:valAx>\n"
            '        <c:axId val="-2129643912"/>\n'
            "        <c:scaling>\n"
            '          <c:orientation val="minMax"/>\n'
            "        </c:scaling>\n"
            '        <c:delete val="0"/>\n'
            '        <c:axPos val="l"/>\n'
            "        <c:majorGridlines/>\n"
            '        <c:numFmt formatCode="General" sourceLinked="1"/>\n'
            '        <c:majorTickMark val="out"/>\n'
            '        <c:minorTickMark val="none"/>\n'
            '        <c:tickLblPos val="nextTo"/>\n'
            '        <c:crossAx val="-2128940872"/>\n'
            '        <c:crosses val="autoZero"/>\n'
            '        <c:crossBetween val="midCat"/>\n'
            "      </c:valAx>\n"
            "    </c:plotArea>\n"
            "    <c:legend>\n"
            '      <c:legendPos val="r"/>\n'
            "      <c:layout/>\n"
            '      <c:overlay val="0"/>\n'
            "    </c:legend>\n"
            '    <c:plotVisOnly val="1"/>\n'
            '    <c:dispBlanksAs val="gap"/>\n'
            '    <c:showDLblsOverMax val="0"/>\n'
            "  </c:chart>\n"
            "  <c:txPr>\n"
            "    <a:bodyPr/>\n"
            "    <a:lstStyle/>\n"
            "    <a:p>\n"
            "      <a:pPr>\n"
            '        <a:defRPr sz="1800"/>\n'
            "      </a:pPr>\n"
            '      <a:endParaRPr lang="en-US"/>\n'
            "    </a:p>\n"
            "  </c:txPr>\n"
            "</c:chartSpace>\n"
        ) % (self._scatterStyle_val, self._ser_xml)
        return xml

    @property
    def _marker_xml(self):
        pass

    @property
    def _scatterStyle_val(self):
        pass

    @property
    def _ser_xml(self):
        pass

    @property
    def _spPr_xml(self):
        pass


class _BubbleChartXmlWriter(_XyChartXmlWriter):
    """
    Provides specialized methods particular to the ``<c:bubbleChart>``
    element.
    """

    @property
    def xml(self):
        xml = (
            "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>\n"
            '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawin'
            'gml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/draw'
            'ingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/off'
            'iceDocument/2006/relationships">\n'
            "  <c:chart>\n"
            '    <c:autoTitleDeleted val="0"/>\n'
            "    <c:plotArea>\n"
            "      <c:layout/>\n"
            "      <c:bubbleChart>\n"
            '        <c:varyColors val="0"/>\n'
            "%s"
            "        <c:dLbls>\n"
            '          <c:showLegendKey val="0"/>\n'
            '          <c:showVal val="0"/>\n'
            '          <c:showCatName val="0"/>\n'
            '          <c:showSerName val="0"/>\n'
            '          <c:showPercent val="0"/>\n'
            '          <c:showBubbleSize val="0"/>\n'
            "        </c:dLbls>\n"
            '        <c:bubbleScale val="100"/>\n'
            '        <c:showNegBubbles val="0"/>\n'
            '        <c:axId val="-2115720072"/>\n'
            '        <c:axId val="-2115723560"/>\n'
            "      </c:bubbleChart>\n"
            "      <c:valAx>\n"
            '        <c:axId val="-2115720072"/>\n'
            "        <c:scaling>\n"
            '          <c:orientation val="minMax"/>\n'
            "        </c:scaling>\n"
            '        <c:delete val="0"/>\n'
            '        <c:axPos val="b"/>\n'
            '        <c:numFmt formatCode="General" sourceLinked="1"/>\n'
            '        <c:majorTickMark val="out"/>\n'
            '        <c:minorTickMark val="none"/>\n'
            '        <c:tickLblPos val="nextTo"/>\n'
            '        <c:crossAx val="-2115723560"/>\n'
            '        <c:crosses val="autoZero"/>\n'
            '        <c:crossBetween val="midCat"/>\n'
            "      </c:valAx>\n"
            "      <c:valAx>\n"
            '        <c:axId val="-2115723560"/>\n'
            "        <c:scaling>\n"
            '          <c:orientation val="minMax"/>\n'
            "        </c:scaling>\n"
            '        <c:delete val="0"/>\n'
            '        <c:axPos val="l"/>\n'
            "        <c:majorGridlines/>\n"
            '        <c:numFmt formatCode="General" sourceLinked="1"/>\n'
            '        <c:majorTickMark val="out"/>\n'
            '        <c:minorTickMark val="none"/>\n'
            '        <c:tickLblPos val="nextTo"/>\n'
            '        <c:crossAx val="-2115720072"/>\n'
            '        <c:crosses val="autoZero"/>\n'
            '        <c:crossBetween val="midCat"/>\n'
            "      </c:valAx>\n"
            "    </c:plotArea>\n"
            "    <c:legend>\n"
            '      <c:legendPos val="r"/>\n'
            "      <c:layout/>\n"
            '      <c:overlay val="0"/>\n'
            "    </c:legend>\n"
            '    <c:plotVisOnly val="1"/>\n'
            '    <c:dispBlanksAs val="gap"/>\n'
            '    <c:showDLblsOverMax val="0"/>\n'
            "  </c:chart>\n"
            "  <c:txPr>\n"
            "    <a:bodyPr/>\n"
            "    <a:lstStyle/>\n"
            "    <a:p>\n"
            "      <a:pPr>\n"
            '        <a:defRPr sz="1800"/>\n'
            "      </a:pPr>\n"
            '      <a:endParaRPr lang="en-US"/>\n'
            "    </a:p>\n"
            "  </c:txPr>\n"
            "</c:chartSpace>\n"
        ) % self._ser_xml
        return xml

    @property
    def _bubble3D_val(self):
        pass

    @property
    def _ser_xml(self):
        pass


class _CategorySeriesXmlWriter(_BaseSeriesXmlWriter):
    """
    Generates XML snippets particular to a category chart series.
    """

    @property
    def cat(self):
        """
        Return the ``<c:cat>`` element XML for this series, as an oxml
        element.
        """
        pass

    @property
    def cat_xml(self):
        """
        The unicode XML snippet for the ``<c:cat>`` element for this series,
        containing the category labels and spreadsheet reference.
        """
        pass

    @property
    def val(self):
        """
        The ``<c:val>`` XML for this series, as an oxml element.
        """
        pass

    @property
    def val_xml(self):
        """
        Return the unicode XML snippet for the ``<c:val>`` element describing
        this series, containing the series values and their spreadsheet range
        reference.
        """
        pass

    @property
    def _cat_num_pt_xml(self):
        """
        The unicode XML snippet for the ``<c:pt>`` elements when category
        labels are numeric (including date type).
        """
        pass

    @property
    def _cat_pt_xml(self):
        """
        The unicode XML snippet for the ``<c:pt>`` elements containing the
        category names for this series.
        """
        pass

    @property
    def _cat_tmpl(self):
        """
        The template for the ``<c:cat>`` element for this series, containing
        the category labels and spreadsheet reference.
        """
        pass

    def _lvl_xml(self, categories):
        """
        The unicode XML snippet for the ``<c:lvl>`` elements containing
        multi-level category names.
        """
        pass

    @property
    def _multiLvl_cat_tmpl(self):
        """
        The template for the ``<c:cat>`` element for this series when there
        are multi-level (nested) categories.
        """
        pass

    @property
    def _numRef_cat_tmpl(self):
        """
        The template for the ``<c:cat>`` element for this series when the
        labels are numeric (or date) values.
        """
        pass

    @property
    def _val_pt_xml(self):
        """
        The unicode XML snippet containing the ``<c:pt>`` elements containing
        the values for this series.
        """
        pass

    @property
    def _val_tmpl(self):
        """
        The template for the ``<c:val>`` element for this series, containing
        the series values and their spreadsheet range reference.
        """
        pass


class _XySeriesXmlWriter(_BaseSeriesXmlWriter):
    """
    Generates XML snippets particular to an XY series.
    """

    @property
    def xVal(self):
        """
        Return the ``<c:xVal>`` element for this series as an oxml element.
        This element contains the X values for this series.
        """
        pass

    @property
    def xVal_xml(self):
        """
        Return the ``<c:xVal>`` element for this series as unicode text. This
        element contains the X values for this series.
        """
        pass

    @property
    def yVal(self):
        """
        Return the ``<c:yVal>`` element for this series as an oxml element.
        This element contains the Y values for this series.
        """
        pass

    @property
    def yVal_xml(self):
        """
        Return the ``<c:yVal>`` element for this series as unicode text. This
        element contains the Y values for this series.
        """
        pass

    @property
    def _xVal_tmpl(self):
        """
        The template for the ``<c:xVal>`` element for this series, containing
        the X values and their spreadsheet range reference.
        """
        pass

    @property
    def _yVal_tmpl(self):
        """
        The template for the ``<c:yVal>`` element for this series, containing
        the Y values and their spreadsheet range reference.
        """
        pass


class _BubbleSeriesXmlWriter(_XySeriesXmlWriter):
    """
    Generates XML snippets particular to a bubble chart series.
    """

    @property
    def bubbleSize(self):
        """
        Return the ``<c:bubbleSize>`` element for this series as an oxml
        element. This element contains the bubble size values for this
        series.
        """
        pass

    @property
    def bubbleSize_xml(self):
        """
        Return the ``<c:bubbleSize>`` element for this series as unicode
        text. This element contains the bubble size values for all the
        data points in the chart.
        """
        pass

    @property
    def _bubbleSize_tmpl(self):
        """
        The template for the ``<c:bubbleSize>`` element for this series,
        containing the bubble size values and their spreadsheet range
        reference.
        """
        pass


class _BubbleSeriesXmlRewriter(_BaseSeriesXmlRewriter):
    """
    A series rewriter suitable for bubble charts.
    """

    def _rewrite_ser_data(self, ser, series_data, date_1904):
        """
        Rewrite the ``<c:tx>``, ``<c:cat>`` and ``<c:val>`` child elements
        of *ser* based on the values in *series_data*.
        """
        pass


class _CategorySeriesXmlRewriter(_BaseSeriesXmlRewriter):
    """
    A series rewriter suitable for category charts.
    """

    def _rewrite_ser_data(self, ser, series_data, date_1904):
        """
        Rewrite the ``<c:tx>``, ``<c:cat>`` and ``<c:val>`` child elements
        of *ser* based on the values in *series_data*.
        """
        pass


class _XySeriesXmlRewriter(_BaseSeriesXmlRewriter):
    """
    A series rewriter suitable for XY (aka. scatter) charts.
    """

    def _rewrite_ser_data(self, ser, series_data, date_1904):
        """
        Rewrite the ``<c:tx>``, ``<c:xVal>`` and ``<c:yVal>`` child elements
        of *ser* based on the values in *series_data*.
        """
        pass
