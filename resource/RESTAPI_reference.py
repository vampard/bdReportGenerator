
import requests

#A class containing API Methods for Black Duck Hub
# For educational purposes only
class HubAPI:

    #When initializing this class it will store the endpoint of
    # the hub instance. Must include the full path.
    #eg http://myhubendpoint.mydomain:8080
    def __init__(self, endpoint):
        self.URL = endpoint
        #Hub uses session authentication.
        #This object will store the session and must be passed to each request
        #to authenticate
        self.aSession = requests.Session()
        ### Hub 4.0 and later all operations are in https
        ## Please work iwth your IT infrastructure to manage HTTP certificates
        self.aSession.verify=False
        requests.packages.urllib3.disable_warnings()
        # Variable to hold CSRF token which must be included in the header of all
        # requests that attempt to modify the system.
        self.CSRF = ''
    #Helper method to build path
    def urlCompose(self, path=''):
        return self.URL + '/' + path

    #Helper method to get Hub API links from Response
    #Params -- self - required pythont parameter
    #          meta  - the meta section of an API Response
    #          tag   -  the name of the link to get
    def getLink(self, meta, tag):
        for i in range(len(meta['links'])):
            if meta['links'][i]['rel'] == tag:
                return meta['links'][i]['href']
        print( "ERROR: getLink(" + tag +') failed. Tag not found')
        return 0
    #Authenticates the session
    def authenticate( self, username, password):
        # Username and password will be sent in body of post request
        authParam = {'j_username':username, 'j_password':password}
        #Send a post request to authentication endpoint
        response = self.aSession.post(self.urlCompose('j_spring_security_check'),
            data = authParam)

        #check for Success
        if response.ok:
            self.CSRF = response.headers['x-csrf-token']
            return 1
        else:
            print("Error in authentication")
            return 0

    #Sends request to project endpoint. Parameters are optional and mapped directly
    # from API documentation.
    def getProjects( self, limit=100, offset=0,sort='',q=''):
        payload = {'limit':limit, 'offset':offset, 'sort':sort, 'q':q}
        response = self.aSession.get(self.urlCompose('api/projects'), params=payload)
        if response.ok:
            return response.json()
        else:
            print('Bad response in getProjects')
            return response.json()

    #Sends request to version endpoint for a given project.
    #This endpoint will be retrived from the project response data and passed
    # directly here as projectURL.
    def getVersions( self, projectURL, limit=100, offset=0, sort='', q=''):
        payload = {'limit':limit, 'offset':offset,'sort':sort, 'q':q}
        response = self.aSession.get(projectURL, params=payload)
        if response.ok:
            return response.json()
        else:
            print('Bad response in getVersions')
            return response.json()

    #Sends post request to have Hub create a report about a version of a project.
    # reportURL comes from the body of the getVersion response.
    def generateReport( self, reportURL):
        reportFormat = {'reportFormat':'CSV'}
        response = self.aSession.post(reportURL, json = reportFormat, headers={'x-csrf-token':self.CSRF})
        if response.ok:
            return response.text
        else:
            print("Error: bad request in generateReport()")
            return response.text

    def getReports( self, reportURL):
        response = self.aSession.get(reportURL)
        if response.ok:
            return response.json()
        else:
            print("Error: bad request in getReports")
    #Downloads the report from reportURL to dest
    def downloadReport(self, reportURL, dest='report.zip'):
        response = self.aSession.get(reportURL)
        with open(dest, 'wb') as output:
            for chunk in response.iter_content(2000):
                output.write(chunk)
