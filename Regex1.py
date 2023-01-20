import re
string = ["GCA_2022_aw_US_BUZZFEED.COM_direct_cld_desk_web_ros_geo_t2-zip-lifestyleent_disp_standrd_300 x 600_NA_na_CPM_3rd_EN_G22GCA01US2_P1ZV26G_20220418_20220717"]

for element in string:
    z = re.search('(3rd\w+)',element)
    if z:
        print((z.groups()))