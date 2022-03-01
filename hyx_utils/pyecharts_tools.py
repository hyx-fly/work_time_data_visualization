import os.path
from pyecharts import options as opts
from pyecharts.charts import Timeline, Page, Pie, Line, Bar
from pyecharts.globals import CurrentConfig
from pyecharts.globals import WarningType
from bs4 import BeautifulSoup


def gen_bar(data, title_str='bar_例子', subtitle_str='副标题', is_show=False, bar_id='bar_0'):
    x_data = [x[0] for x in data]
    y_data = [y[1] for y in data]
    bar = (
        Bar()
        .add_xaxis(x_data)
        .add_yaxis('', y_data)
        .set_global_opts(title_opts=opts.TitleOpts(title=title_str, subtitle=subtitle_str),
                         datazoom_opts=opts.DataZoomOpts(is_show=True,
                                                         range_start=0,
                                                         range_end=100,
                                                         pos_top='5%'),
                         visualmap_opts=opts.VisualMapOpts(max_=250,
                                                           min_=0,
                                                           is_show=True,
                                                           pos_top='30%',
                                                           pos_left='1.5%'),
                         toolbox_opts=opts.ToolboxOpts(is_show=True,
                                                       orient='vertical',
                                                       pos_left='5%',
                                                       pos_bottom='12%'),
                         legend_opts=opts.LegendOpts(is_show=True)
                         )
    )
    bar.chart_id = bar_id
    if is_show:
        bar.render('bar_example.html')

    return bar


def gen_line(data, title_str='line_例子', subtitle_str='副标题', is_show=False, line_id='line_0'):
    x_data = [f'{x[0]}日' for x in data]
    y_data = [y[1] for y in data]
    line = (
        Line()
        .add_xaxis(x_data)
        .add_yaxis('', y_data,
                   is_smooth=True,  # 曲线平滑
                   markline_opts=opts.MarkLineOpts(data=[opts.MarkLineItem(type_='average')]))  # 显示平均值
        .set_global_opts(title_opts=opts.TitleOpts(title=title_str, subtitle=subtitle_str))
    )
    line.chart_id = line_id
    if is_show:
        line.render('line_example.html')

    return line


def gen_pie(data, title_str='pie_例子', subtitle_str='副标题', is_show=False, pie_id='pie_0'):
    pie = (
        Pie()
        .add('', data,
             radius=['30%', '75%'],
             rosetype='radius')
        .set_global_opts(title_opts=opts.TitleOpts(title=title_str, subtitle=subtitle_str))
        .set_series_opts(label_opts=opts.LabelOpts(formatter='{b}: {d}%'))
    )
    pie.chart_id = pie_id
    if is_show:
        pie.render('pie_example.html')

    return pie


def modify_html_background(html_path, attribute_text):
    with open(html_path, 'r', encoding='UTF-8') as file:
        soup = BeautifulSoup(file, 'lxml')
        tag = soup.body
        tag['background'] = attribute_text

    with open(html_path, 'wb') as new_file:
        new_file.write(soup.prettify('UTF-8'))


def gen_timeline(is_show=False, play_interval=1500, tl_id='timeline_0'):
    tl = Timeline()
    tl.add_schema(is_auto_play=True, play_interval=play_interval)
    tl.chart_id = 'data_fx_timeline'
    if is_show:
        tl.render('tl_example.html')
    tl.chart_id = tl_id

    return tl


def gen_page_html(charts, title_str='示例文件', path='example.html'):
    CurrentConfig.ONLINE_HOST = './pyecharts-assets-master/assets/'
    # 关闭警告提示
    WarningType.ShowWarning = False
    #  用于整合所有图表的page
    page = Page(layout=Page.DraggablePageLayout, page_title=title_str)
    if type(charts) is list or type(charts) is tuple:
        for chart in charts:
            page.add(chart)
    else:
        page.add(charts)
    page.render(path)
    chart_config_path = './data/output/pyecharts-assets-master/chart_config.json'
    if os.path.exists(chart_config_path):
        page.save_resize_html(source=path, cfg_file=chart_config_path, dest=path)

    modify_html_background(path, './pyecharts-assets-master/log.png')
    return path
