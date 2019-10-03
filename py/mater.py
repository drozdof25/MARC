import os
from py_mentat import*

def add_experiment_table(directory,file):
    file_tbl = directory+'\\'+file
    py_send('@popup(table_pm,0) *md_table_read_any "%s"' % file_tbl)
    file = file.replace('.txt','')
    file = file.replace(' ','')
    py_send('*edit_table %s' % file)
    py_send('*set_md_table_type 1')
    py_send('experimental_data')
    py_send('*xcv_table uniaxial')
    py_send(file)
    py_send('*xcv_model mooney2')
    py_send('*xcv_mode uniaxial on')
    py_send('*xcv_positive on')
    py_send('*xcv_extrapolation on')
    py_send('*xcv_compute')
    py_send('*xcv_create')
    mater_name = py_get_string('mater_name()')
    py_send('*edit_mater %s' % mater_name)
    py_send('*mater_name %s' % file)
data_directory  = 'D:\\buffer\\data'
files = os.listdir(data_directory)
for file in files:
    add_experiment_table(data_directory,file)

