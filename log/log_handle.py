import re
import xlwt
from datetime import datetime
from operator import itemgetter

def action ():
    action_dict={}
    with open('case_log.txt',mode='r',encoding='utf-8') as file:
	    lines = file.readline()
	    for line in lines:
	#mat=re.search(r'ACTION\s(\w+)\s\[finished\].*duration\=((\d+)\.(\d){3})sec',line)
	        action_start = re.search(r'(\d{1,2}/\d{1,2}\s\d{1,2}:\d{1,2}:\d{1,2},\d{3}\sACTION\s(\w+)\s\[started\]',line)
				if action_start:
	                time_day='/' join(['action_start.group(1)','action_start.group(2)'])
				    time_time=':' join(['action_start.group(3)','action_start.group(4)','action_start.group(5)'])
					time_all=',' join([time_time,'action_start.group(6)'])
					time_start='' join(time_day,time_all)
	                action_name=action_start.group(7)	 
	        action_finished = re.search(r'(\d{1,2}/\d{1,2}\s\d{1,2}:\d{1,2}:\d{1,2},\d{3}\sACTION\s(\w+)\s\[finished\]',line)
			    if action_finised :
	                time_day='/' join(['action_finished.group(1)','action_finished.group(2)'])
				    time_time=':' join(['action_finished.group(3)','action_finished.group(4)','action_finished.group(5)'])
					time_all=',' join([time_time,'action_finished.group(6)'])
					time_finised='' join(time_day,time_all)
	                action_name=action_start.group(7)
		    start_time=datetime.datetime.strptime(time_start,'%m/%d %H:%M:%S,%f')
			finished_time=datetime.datetime.strptime(time_finished,'%m/%d %H:%M:%S,%f')
			duration=(finished_time-start_time).seconds
			    if action_name not in action_dict:
                    name_dict={}
                    name_dict['count'] = 1
                    name_dict['duration'] = float(duration)
                    action_dict[action_name] = name_dict
                else:
                    action_dict[action_name]['count'] = action_dict[action_name]['count'] + 1
                    action_dict[action_name]['duration'] = float(action_dict[action_name]['duration']) + float(duration)
	
all_list=[]
each_tuple=()
for key in sorted(action_dict):
    #print (key + ' ----------- ' + str(action_dict[key]['count'])+ ' -------- ' + str(action_dict[key]['duration']))
    action_dict[key]['avarage'] = action_dict[key]['duration']/action_dict[key]['count']
    #print (key + ' ----------- ' + str(action_dict[key]['count'])+ ' -------- ' + str(action_dict[key]['avarage']))
    each_action_list=[]
    each_action_list.append(key)
    each_action_list.append(round(action_dict[key]['duration'],3))
    each_action_list.append(action_dict[key]['count'])
    each_action_list.append(round(action_dict[key]['avarage'],3))
    each_tuple=tuple(each_action_list)
    all_list.append(each_tuple)
all_list.sort(key=itemgetter(2))

print ('#------------------------------------- The next will be write data to excel ----------------------------------------------#')

def set_style(dStyle):
    font = xlwt.Font()
    if 'blod_mode'in dStyle and dStyle['blod_mode'] == 'enable':
        font.bold = 'on'
    font.colour_index = xlwt.Style.colour_map[dStyle['font_color']]
    pattern = xlwt.Pattern()
    if 'background_color' in dStyle:
        pattern.pattern = xlwt.Pattern.SOLID_PATTERN
        pattern.pattern_fore_colour = xlwt.Style.colour_map[dStyle['background_color']]
    #Setting the Alignment for the Contents of a Cell
    alignment = xlwt.Alignment()
    if 'h_align_mode' in dStyle:
        alignment.horz = xlwt.Alignment.HORZ_CENTER
    if 'v_align_mode' in dStyle:
        alignment.vert = xlwt.Alignment.VERT_CENTER
    
    style = xlwt.XFStyle() 
    style.pattern = pattern
    style.font = font
    style.alignment = alignment
    return style

d_style1={'font_color':'black','background_color':'blue','blod_mode':'enable'}
d_style2={'font_color':'black','background_color':'yellow'}
d_style3={'font_color':'black','background_color':'yellow','h_align_mode':'HORZ_CENTER','v_align_mode':'VERT_CENTER'}
d_style4={'font_color':'black','h_align_mode':'HORZ_CENTER','v_align_mode':'VERT_CENTER'}
my_style1=set_style(d_style1)
my_style2=set_style(d_style2)
my_style3=set_style(d_style3)
my_style4=set_style(d_style4)

wb = xlwt.Workbook()
sh = wb.add_sheet('action_statistics',cell_overwrite_ok=True)
## Add title for this sheet
sh.write(0,0,'Action Name',my_style1)
sh.write(0,1,'Total Time',my_style1)
sh.write(0,2,'Action Count',my_style1)
sh.write(0,3,'Average Time',my_style1)
sh.col(0).width = 5555
sh.col(1).width = 5555
sh.col(2).width = 5555
sh.col(3).width = 5555
for i,item in enumerate(all_list):
    print (str(i) + ' -----> ' + str(item))
    
    if item[3]> 2:
        my_style=d_style2
        sh.write(i+1,0,item[0],my_style2)
        sh.write(i+1,1,item[1],my_style3)
        sh.write(i+1,2,item[2],my_style3)
        sh.write(i+1,3,item[3],my_style3)
    else:
        sh.write(i+1,0,item[0])
        sh.write(i+1,1,item[1],my_style4)
        sh.write(i+1,2,item[2],my_style4)
        sh.write(i+1,3,item[3],my_style4)

wb.save('my_case_log_analyze.xls')
