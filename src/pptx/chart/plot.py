"""Plot-related objects.

A plot is known as a chart group in the MS API. A chart can have more than one plot overlayed on
each other, such as a line plot layered over a bar plot.
"""

from __future__ import annotations

from pptx.chart.category import Categories
from pptx.chart.datalabel import DataLabels
from pptx.chart.series import SeriesCollection
from pptx.enum.chart import XL_CHART_TYPE as XL
from pptx.oxml.ns import qn
from pptx.oxml.simpletypes import ST_BarDir, ST_Grouping
from pptx.util import lazyproperty


class _BasePlot(object):
    """
    A distinct plot that appears in the plot area of a chart. A chart may
    have more than one plot, in which case they appear as superimposed
    layers, such as a line plot appearing on top of a bar chart.
    """

    def __init__(self, xChart, chart):
        super(_BasePlot, self).__init__()
        self._element = xChart
        self._chart = chart

    @lazyproperty
    def categories(self):
        """
        Returns a |category.Categories| sequence object containing
        a |category.Category| object for each of the category labels
        associated with this plot. The |category.Category| class derives from
        ``str``, so the returned value can be treated as a simple sequence of
        strings for the common case where all you need is the labels in the
        order they appear on the chart. |category.Categories| provides
        additional properties for dealing with hierarchical categories when
        required.
        """
        pass

    @property
    def chart(self):
        """
        The |Chart| object containing this plot.
        """
        pass

    @property
    def data_labels(self):
        """
        |DataLabels| instance providing properties and methods on the
        collection of data labels associated with this plot.
        """
        pass

    @property
    def has_data_labels(self):
        """
        Read/write boolean, |True| if the series has data labels. Assigning
        |True| causes data labels to be added to the plot. Assigning False
        removes any existing data labels.
        """
        pass

    @has_data_labels.setter
    def has_data_labels(self, value):
        """
        Add, remove, or leave alone the ``<c:dLbls>`` child element depending
        on current state and assigned *value*. If *value* is |True| and no
        ``<c:dLbls>`` element is present, a new default element is added with
        default child elements and settings. When |False|, any existing dLbls
        element is removed.
        """
        pass

    @lazyproperty
    def series(self):
        """
        A sequence of |Series| objects representing the series in this plot,
        in the order they appear in the plot.
        """
        pass

    @property
    def vary_by_categories(self):
        """
        Read/write boolean value specifying whether to use a different color
        for each of the points in this plot. Only effective when there is
        a single series; PowerPoint automatically varies color by series when
        more than one series is present.
        """
        pass

    @vary_by_categories.setter
    def vary_by_categories(self, value):
        pass


class AreaPlot(_BasePlot):
    """
    An area plot.
    """


class Area3DPlot(_BasePlot):
    """
    A 3-dimensional area plot.
    """


class BarPlot(_BasePlot):
    """
    A bar chart-style plot.
    """

    @property
    def gap_width(self):
        """
        Width of gap between bar(s) of each category, as an integer
        percentage of the bar width. The default value for a new bar chart is
        150, representing 150% or 1.5 times the width of a single bar.
        """
        pass

    @gap_width.setter
    def gap_width(self, value):
        pass

    @property
    def overlap(self):
        """
        Read/write int value in range -100..100 specifying a percentage of
        the bar width by which to overlap adjacent bars in a multi-series bar
        chart. Default is 0. A setting of -100 creates a gap of a full bar
        width and a setting of 100 causes all the bars in a category to be
        superimposed. A stacked bar plot has overlap of 100 by default.
        """
        pass

    @overlap.setter
    def overlap(self, value):
        """
        Set the value of the ``<c:overlap>`` child element to *int_value*,
        or remove the overlap element if *int_value* is 0.
        """
        pass


class BubblePlot(_BasePlot):
    """
    A bubble chart plot.
    """

    @property
    def bubble_scale(self):
        """
        An integer between 0 and 300 inclusive indicating the percentage of
        the default size at which bubbles should be displayed. Assigning
        |None| produces the same behavior as assigning `100`.
        """
        pass

    @bubble_scale.setter
    def bubble_scale(self, value):
        pass


class DoughnutPlot(_BasePlot):
    """
    An doughnut plot.
    """


class LinePlot(_BasePlot):
    """
    A line chart-style plot.
    """


class PiePlot(_BasePlot):
    """
    A pie chart-style plot.
    """


class RadarPlot(_BasePlot):
    """
    A radar-style plot.
    """


class XyPlot(_BasePlot):
    """
    An XY (scatter) plot.
    """


def PlotFactory(xChart, chart):
    """
    Return an instance of the appropriate subclass of _BasePlot based on the
    tagname of *xChart*.
    """
    pass


class PlotTypeInspector(object):
    """
    "One-shot" service object that knows how to identify the type of a plot
    as a member of the XL_CHART_TYPE enumeration.
    """

    @classmethod
    def chart_type(cls, plot):
        """
        Return the member of :ref:`XlChartType` that corresponds to the chart
        type of *plot*.
        """
        pass

    @classmethod
    def _differentiate_area_3d_chart_type(cls, plot):
        pass

    @classmethod
    def _differentiate_area_chart_type(cls, plot):
        pass

    @classmethod
    def _differentiate_bar_chart_type(cls, plot):
        pass

    @classmethod
    def _differentiate_bubble_chart_type(cls, plot):
        pass

    @classmethod
    def _differentiate_doughnut_chart_type(cls, plot):
        pass

    @classmethod
    def _differentiate_line_chart_type(cls, plot):
        pass

    @classmethod
    def _differentiate_pie_chart_type(cls, plot):
        pass

    @classmethod
    def _differentiate_radar_chart_type(cls, plot):
        pass

    @classmethod
    def _differentiate_xy_chart_type(cls, plot):
        pass
