Attribute VB_Name = "Module1"
Public Const HTTPGETRANGE = "GET {PATH} HTTP/1.1" & vbCrLf & _
                            "Range: bytes={RANGE}" & vbCrLf & _
                            "Host: {HOST}" & vbCrLf & _
                            "Accept: */*" & vbCrLf & _
                            "User-Agent: Simple Download Manager v0.1" & vbCrLf & _
                            vbCrLf

Public Const HTTPGET = "GET {PATH} HTTP/1.1" & vbCrLf & _
                       "Host: {HOST}" & vbCrLf & _
                       "Accept: */*" & vbCrLf & _
                       "User-Agent: Simple Download Manager v0.1" & vbCrLf & _
                       vbCrLf

Public Const title = ".:[SDM]:.  coded by chown"
