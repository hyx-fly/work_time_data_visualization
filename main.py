# coding:utf-8
from hyx_utils.hyx_tools import *
from hyx_utils.pyecharts_tools import gen_bar, gen_pie, gen_page_html, gen_line, gen_timeline


def print_version(version):
    print(f'欢迎使用《乘务劳时数据可视化系统》 version：{version}')


@print_execute_time
def app_main(out_date, gen_chart=True):
    time_out(out_date)
    data_path, mon_num, today_num = gen_work_time_main()
    # if today_num == 1:
    #     today_num = 32

    if gen_chart is False:
        open_xlsx(data_path, is_show=True)

    elif gen_chart:
        open_xlsx(data_path, is_show=False)
        bar_data = read_total_work_time(data_path, '统计', 'AI')
        bar_data.sort(key=lambda x: x[1], reverse=True)
        pie_data = work_time_type_count(bar_data)
        line_data = read_total_work_time_one_day(data_path, today_num, '日劳时统计')

        # 生成静态柱状图
        bar_static_0 = gen_bar(data=bar_data,
                               title_str=f'{mon_num}月{today_num - 1}日-个人劳时',
                               subtitle_str='(单位: 小时)',
                               is_show=False,
                               bar_id='bar_static_0')

        # 生成静态饼状图
        pie_static_0 = gen_pie(data=pie_data,
                               title_str=f'{mon_num}月{today_num - 1}日-劳时分布',
                               subtitle_str='占比',
                               is_show=False)

        # 生成静态折线图
        line_static_0 = gen_line(data=line_data,
                                 title_str=f'{mon_num}月1日-{today_num - 1}日-劳时波动',
                                 subtitle_str='(单位: 小时)',
                                 is_show=False)

        # 生成时间轴柱状图
        data_col = ('D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T',
                    'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH')
        bar_timeline = gen_timeline(is_show=False, play_interval=200)
        stop_num = today_num
        # stop_num = 32
        for col in data_col:
            date_num = data_col.index(col) + 1
            if date_num == stop_num:
                break
            tl_bar_data = read_total_work_time(data_path, '累积劳时', col)
            tl_bar_data.sort(key=lambda x: x[1], reverse=True)
            bar = gen_bar(data=tl_bar_data,
                          title_str=f'{mon_num}月{date_num}日-个人累积劳时',
                          subtitle_str='(单位: 小时)',
                          is_show=False,
                          bar_id='bar_dynamic_0')
            bar_timeline.add(chart=bar, time_point=f'{date_num}日')

        charts_list = [bar_static_0, pie_static_0, line_static_0, bar_timeline]

        open_html(gen_page_html(charts_list, title_str='劳时分析', path='./data/output/劳时分析.html'))


if __name__ == '__main__':
    print('第二个示例')
    print_version('2022.02.01_1')
    try:
        app_main('2023/02/01')
    except BaseException as msg:
        hyx_log(f'程序执行时发生错误，错误信息:{msg}')
