VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vbWebMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private pIE As Object      'Internet Explorer
Private document As Object 'Document object
Private window As Object   'Window object
Private pLeaflet As Object  'Leaflet
Private pAPI As Object      'Site API

Public Enum EMarkerSize
  SizeSmall
  SizeMedium
  SizeLarge
End Enum
Public Enum EMarkerSymbol
  symCircleStroked
  symCircle
  symSquareStroked
  symSquare
  symTriangleStroked
  symTriangle
  symStarStroked
  symStar
  symCross
  symMarkerStroked
  symMarker
  symReligiousJewish
  symReligiousChristian
  symReligiousMuslim
  symCemetry
  symRocket
  symAirport
  symHeliport
  symRail
  symRailMetro
End Enum


Public Function Create() As vbWebMap
  Set Create = New vbWebMap
  Call Create.protInit
End Function


Public Property Get API() As Object
  Set API = pAPI
End Property
Public Property Get Leaflet() As Object
  Set Leaflet = pLeaflet
End Property
Public Sub panTo(ByVal Lat As Double, ByVal Lng As Double)
  Call document.panTo(Lat, Lng)
End Sub
Public Sub setBounds(ByVal Lat1 As Double, ByVal Lng1 As Double, ByVal Lat2 As Double, ByVal Lng2 As Double)
  Call document.setBounds(Lat1, Lng1, Lat2, Lng2)
End Sub
Public Sub addLayer(ByVal layer As gisFeatureCollection)
  Dim s As String: s = layer.ToString()
  Call document.addLayer(s)
  Dim vBounds: vBounds = layer.Bounds
  Call setBounds(vBounds(0), vBounds(1), vBounds(2), vBounds(3))
End Sub

Public Function getMarkerStyle(ByVal color As Long, sym As EMarkerSymbol, size As EMarkerSize) As Object
  Dim o As Object: Set o = CreateObject("Scripting.Dictinary")
  o("marker-color") = "#" & Hex(color)
  o("marker-size") = Switch(size, SizeLarge, "large", SizeMedium, "medium", SizeSmall, "small")
End Function

Public Property Get isIERunning() As Boolean
  On Error GoTo ErrorOccurred
  isIERunning = pIE.Type = "HTML Document"
ErrorOccurred:
End Property





Friend Sub protInit()
  Set pIE = CreateObject("InternetExplorer.Application")
  pIE.MenuBar = False
  pIE.AddressBar = False
  pIE.Resizable = True
  pIE.StatusBar = False

  pIE.Visible = True
  Call pIE.Navigate("https://geojson.io/#map=2/20.0/0.0")
  Do While pIE.ReadyState = 4: DoEvents: Loop   'Wait for loading status
  Do Until pIE.ReadyState = 4: DoEvents: Loop   'Wait for ready status
  Set document = pIE.document
  Set window = document.parentWindow
  
  'Bind API
  Call AppendJS("document.api = window.api;" & _
                "document.L   = window.L;" & _
                "document.querySelector(""#geojsonio-body > div > div > div.right > div.top > button"").click();" & _
                "document.panTo = function(lat,lng){window.api.map.panTo(new L.LatLng(lat, lng))};" & _
                "document.setBounds = function(lat1,lng1,lat2,lng2){window.api.map.fitBounds([[lat1,lng1],[lat2,lng2]])};" & _
                "document.addLayer = function(s){window.api.data.set({ map: JSON.parse(s)}) };")
  
  Set pAPI = document.API
  Set pLeaflet = document.L
End Sub
Private Sub AppendJS(ByVal sScript As String)
  Set sc = document.createElement("script")
  sc.innerHTML = sScript
  Call document.body.appendChild(sc)
End Sub
Private Sub Class_Terminate()
  Call pIE.Quit
  Set pIE = Nothing
End Sub


Private Function MarkerSymToString(ByVal sym As EMarkerSymbol) As String
  Select Case sym
    Case symCircleStroked: MarkerSymToString = "circle-stroked"
    Case symCircle: MarkerSymToString = "circle"
    Case symSquareStroked: MarkerSymToString = "square-stroked"
    Case symSquare: MarkerSymToString = "square"
    Case symTriangleStroked: MarkerSymToString = "triangle-stroked"
    Case symTriangle: MarkerSymToString = "triangle"
    Case symStarStroked: MarkerSymToString = "star-stroked"
    Case symStar: MarkerSymToString = "star"
    Case symCross: MarkerSymToString = "cross"
    Case symMarkerStroked: MarkerSymToString = "marker-stroked"
    Case symMarker: MarkerSymToString = "marker"
    Case symReligiousJewish: MarkerSymToString = "religious-jewish"
    Case symReligiousChristian: MarkerSymToString = "religious-christian"
    Case symReligiousMuslim: MarkerSymToString = "religious-muslim"
    Case symCemetry: MarkerSymToString = "cemetry"
    Case symRocket: MarkerSymToString = "rocket"
    Case symAirport: MarkerSymToString = "airport"
    Case symHeliport: MarkerSymToString = "heliport"
    Case symRail: MarkerSymToString = "rail"
    Case symRailMetro: MarkerSymToString = "rail-metro"
  End Select
End Function

