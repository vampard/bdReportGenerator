##########################################################################################
 
# Certain source files written and distributed by Leith Jun are subject to the AGPLv3.
# Otherwise you should ask me for the permission to use this package.

# Author : Leith Jun
# Contributor : Jihye Kwon
# Email : leithjun@osbc.co.kr

#    Copyright OSBC Inc. 2019  Leith Jun

#    This program is free software: you can redistribute it and/or modify
#    it under the terms of the GNU Affero General Public License as
#    published by the Free Software Foundation, either version 3 of the
#    License, or (at your option) any later version.

#    This program is distributed in the hope that it will be useful,
#    but WITHOUT ANY WARRANTY; without even the implied warranty of
#    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#    GNU Affero General Public License for more details.

#    You should have received a copy of the GNU Affero General Public License
#    along with this program.  If not, see <https://www.gnu.org/licenses/>.
 

##########################################################################################

import os
import csv
import requests
import sys
import logging
import json

from bdrpkg.reportGenerator import blackduckRPT




if __name__ == "__main__":

    logging.basicConfig(level=logging.DEBUG)
    #logging.disable(logging.DEBUG)
    '''
    arg = ','.join(sys.argv[1:])
    arg = arg.split(",")
    for i in range(0,5,2):
        if arg[i] == "--http" or arg[i] == "--user" or arg[i] == "--password":
            logging.debug("all clear!! with parameters")

        else:
            print(arg[i])
            sys.exit('wrong parameter!')
    br = blackduckRPT(arg[1],arg[3],arg[5])
    '''
    bd_domain = "http://myhubendpoint.mydomain:8080"
    userid = "userid"
    password = "password"
    br = blackduckRPT(bd_domain,userid,password)
    # 이걸로 하면 전체 리포트 산출
    tempInfo = br.getProjectsAndVersions()
    projectInfo = json.loads(tempInfo)
    
    
    for i in projectInfo["KEPCO"]:
        br = blackduckRPT(bd_domain, userid, password)
        myProject = br.findIdentity(i["projectName"],i["versionName"])
        br.createExcel(myProject)




