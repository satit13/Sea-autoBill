#Define clAsync .F.
#Define ccAPIKey '{your API key}'
loRequest = Createobject('MsXml2.XmlHttp')
loRequest.SetRequestHeader('X-CustomerToken',ccAPIKey)
&&loRequest.Open("GET","mydomain.co.uk/Authentication/Authenticate", clAsync)
loRequest.Open("GET","http://192.168.0.89:8080/SmartQWs/branch", clAsync)

loRequest.Send()
? loRequest.Status
? loRequest.ResponseText 