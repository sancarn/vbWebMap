VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "gisOSGB36"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Implements gisIPoint

'Ordnance Survey Grid Reference utility
'Description:
'  Formulation implemented here due to Thomas, Redfearn, etc is as published by OS, but is inferior
'  Ordnance Survey OSGB grid references provide geocoordinate references for UK mapping purposes
'    www.ordnancesurvey.co.uk/docs/support/guide-coordinate-systems-great-britain.pdf
'  Note OSGB grid references cover Great Britain only; Ireland and the Channel Islands have their
'  own references
'
'  Note that these formulae are based on ellipsoidal calculations, and according to the OS are
'  accurate to about 4–5 metres – for greater accuracy, a geoid-based transformation (OSTN15) must
'  be used

'Version History:
'  * 2015-2019  Javascript rendition by [Chris Veness](www.movable-type.co.uk/scripts/latlong-gridref.html)
'  * 2021-05-16 Intiial VBA version by [Sancarn](http://github.com/sancarn)


Public Enum EGoogleMapsBaseMaps
    GoogleBaseMapRoadMap
    GoogleBaseMapSatelite
    GoogleBaseMapTerrain
End Enum



'Note that these formulae are based on ellipsoidal calculations, and according to the OS are
'accurate to about 4–5 metres – for greater accuracy, a geoid-based transformation (OSTN15) must
'be used.
'
'Note Original was converted in 2015 to work with WGS84 by default, OSGB36 as option;
'  www.ordnancesurvey.co.uk/blog/2014/12/confirmation-on-changes-to-latitude-and-longitude
'that algorithm is not used here as it involves use of LatLonEllipsoidal class which isn't implemented yet)




'/* OsGridRef  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - */




Private pEastings As Double
Private pNorthings As Double

Private pLatLngAv As Boolean
Private pLatitude As Double
Private pLongitude As Double

'Constructors
Public Function Create(ByVal Eastings As Double, ByVal Northings As Double) As gisOSGB36
  Set Create = New gisOSGB36
  Call Create.protInit(Eastings, Northings)
End Function
Public Function CreateFromLatLng(ByVal fLatitude As Double, ByVal fLongitude As Double) As gisOSGB36
  Dim vEN As Variant: vEN = getEastingsNorthings(fLatitude, fLongitude)
  Set CreateFromLatLng = New gisOSGB36
  Call CreateFromLatLng.protInit(vEN(0), vEN(1), fLatitude, fLongitude)
End Function
Public Function CreateFromString(ByVal sGridReference As String) As gisOSGB36
  sGridReference = Trim(sGridReference)
  
  Dim oMatch As Object
  Static rxDigString As Object: If rxDigString Is Nothing Then Set rxDigString = regex("^(\d+)\s*[, ;-]\s*(\d+)$")
  If rxDigString.Test(sGridReference) Then
    Set oMatch = rxDigString.Execute(sGridReference)
    Set CreateFromString = Create(oMatch(0).submatches(0), oMatch(0).submatches(1))
    Exit Function
  End If
  
  Static rxGridRef10 As Object: If rxGridRef10 Is Nothing Then Set rxGridRef10 = regex("^([A-Z]{2})\s*([0-9]{5})\s*([0-9]{5})$", "i")
  Static rxGridRef6 As Object: If rxGridRef6 Is Nothing Then Set rxGridRef6 = regex("^([A-Z]{2})\s*([0-9]{3})\s*([0-9]{3})$", "i")
  Dim iDig As Long: iDig = 0
  Select Case True
    Case rxGridRef10.Test(sGridReference)
      iDig = 10
    Case rxGridRef6.Test(sGridReference)
      iDig = 6
  End Select
  
  If iDig > 0 Then
    'Get matched regex
    If iDig = 10 Then
      Set oMatch = rxGridRef10.Execute(sGridReference)
    Else
      Set oMatch = rxGridRef6.Execute(sGridReference)
    End If
    
    'First 2 letters --> e100km and n100km
    Dim sChars As String: sChars = oMatch(0).submatches(0)
    
    'get numeric values of letter references, mapping A->0, B->1, C->2, etc:
    Dim i1 As Long: i1 = Asc(UCase(Mid(sChars, 1, 1))) - Asc("A")
    Dim i2 As Long: i2 = Asc(UCase(Mid(sChars, 2, 1))) - Asc("A")
    
    'shuffle down letters after 'I' since 'I' is not used in grid:
    If i1 > 7 Then i1 = i1 - 1
    If i2 > 7 Then i2 = i2 - 1
    
    'sanity check
    If (i1 < 8 Or i1 > 18) Then Error.Raise 1, "gisOSGB36::CreateFromString", "Invalid grid reference """ & sGridReference & """"
    
    'convert grid letters into 100km-square indexes from false origin (grid square SV):
    Dim e100km As Double: e100km = ((i1 - 2) Mod 5) * 5 + (i2 Mod 5)
    Dim n100km As Double: n100km = (19 - floor(i1 / 5) * 5) - floor(i2 / 5)
    
    'Find easting and northing coords
    Dim fEasting As Long, fNorthing As Long
    If iDig = 10 Then
      fEasting = e100km & oMatch(0).submatches(1)
      fNorthing = n100km & oMatch(0).submatches(2)
    Else
      fEasting = CLng(e100km & oMatch(0).submatches(1) & "00")
      fNorthing = CLng(n100km & oMatch(0).submatches(2) & "00")
    End If
    
    'Create point
    Set CreateFromString = Create(fEasting, fNorthing)
  Else
    Err.Raise 1, "gisOSGB36::CreateFromString", "Invalid grid reference """ & sGridReference & """"
  End If
End Function

Public Property Get Eastings()
  Eastings = pEastings
End Property
Public Property Get Northings()
  Northings = pNorthings
End Property
Public Property Get Latitude()
  If Not pLatLngAv Then Call initLatLng
  Latitude = pLatitude
End Property
Public Property Get Longitude()
  If Not pLatLngAv Then Call initLatLng
  Longitude = pLongitude
End Property



'Obtaining URL's for various systems
Public Property Get URLGoogleEarth(Optional ByVal iScaling As Long = 100000) As String
    URLGoogleEarth = "https://earth.google.com/web/@" & Latitude & "," & Longitude & "," & iScaling & "d"
End Property
Public Property Get URLGoogleMaps(Optional ByVal iZoom As Long = 17, Optional iBaseMap As EGoogleMapsBaseMaps = EGoogleMapsBaseMaps.GoogleBaseMapRoadMap) As String
    Dim sBaseMap As String
    Select Case iBaseMap
        Case EGoogleMapsBaseMaps.GoogleBaseMapRoadMap:  sBaseMap = "roadmap"
        Case EGoogleMapsBaseMaps.GoogleBaseMapSatelite: sBaseMap = "satellite"
        Case EGoogleMapsBaseMaps.GoogleBaseMapTerrain:  sBaseMap = "terrain"
    End Select
    
    URLGoogleMaps = "https://www.google.com/maps/@?api=1&map_action=map&center=" & Latitude & "," & Longitude & "&zoom=" & iZoom & "&basemap=" & sBaseMap
End Property
Public Property Get URLGoogleStreetView() As String
    URLGoogleStreetView = "https://www.google.com/maps/@?api=1&map_action=pano&viewpoint=" & Latitude & "," & Longitude
End Property






Friend Sub protInit(ByVal iEastings As Double, ByVal iNorthings As Double, Optional ByVal Lat, Optional ByVal Lng)
  pEastings = iEastings
  pNorthings = iNorthings
  If Not IsMissing(Lat) Then
    pLatLngAv = True
    pLatitude = Lat
    pLongitude = Lng
  End If
End Sub

'Converts latitude/longitude to Ordnance Survey grid reference easting/northing coordinate.
Private Function getEastingsNorthings(fLatitude As Double, fLongitude As Double) As Variant
  Dim phi As Double: phi = toRad(fLatitude)
  Dim lambda As Double: lambda = toRad(fLongitude)
  
  'Airy 1830 major & minor semi-axes
  Const a As Double = 6377563.396, b As Double = 6356256.909
  
  'NatGrid scale factor on central meridian
  Const F0 As Double = 0.9996012717

  'NatGrid true origin is 49°N,2°W
  Dim phi0 As Double: phi0 = toRad(49)
  Dim lambda0 As Double: lambda0 = toRad(-2)

  'northing & easting of true origin, metres
  Const N0 As Double = -100000, E0 As Double = 400000

  'eccentricity squared
  Dim e2 As Double: e2 = 1 - (b * b) / (a * a)

  'n, n², n³
  Dim n As Double: n = (a - b) / (a + b)
  Dim n2 As Double: n2 = n * n
  Dim n3 As Double: n3 = n2 * n

  'Caching
  Dim cosphi As Double: cosphi = Cos(phi)
  Dim sinphi As Double: sinphi = Sin(phi)

  Dim nu As Double: nu = a * F0 / Sqr(1 - e2 * sinphi * sinphi)              'nu = transverse radius of curvature
  Dim rho As Double: rho = a * F0 * (1 - e2) / ((1 - e2 * sinphi * sinphi) ^ 1.5)  'rho = meridional radius of curvature
  Dim eta2 As Double: eta2 = nu / rho - 1                                          'eta = ?
  
  
  Dim Ma As Double: Ma = (1 + n + (5 / 4) * n2 + (5 / 4) * n3) * (phi - phi0)
  Dim Mb As Double: Mb = (3 * n + 3 * n * n + (21 / 8) * n3) * Sin(phi - phi0) * Cos(phi + phi0)
  Dim Mc As Double: Mc = ((15 / 8) * n2 + (15 / 8) * n3) * Sin(2 * (phi - phi0)) * Cos(2 * (phi + phi0))
  Dim Md As Double: Md = (35 / 24) * n3 * Sin(3 * (phi - phi0)) * Cos(3 * (phi + phi0))
  
  'meridional arc
  Dim M As Double: M = b * F0 * (Ma - Mb + Mc - Md)
  
  Dim cos3phi As Double: cos3phi = cosphi * cosphi * cosphi
  Dim cos5phi As Double: cos5phi = cos3phi * cosphi * cosphi
  Dim tan2phi As Double: tan2phi = Tan(phi) * Tan(phi)
  Dim tan4phi As Double: tan4phi = tan2phi * tan2phi
  Dim i As Double: i = M + N0
  Dim II As Double: II = (nu / 2) * sinphi * cosphi
  Dim III As Double: III = (nu / 24) * sinphi * cos3phi * (5 - tan2phi + 9 * eta2)
  Dim IIIA As Double: IIIA = (nu / 720) * sinphi * cos5phi * (61 - 58 * tan2phi + tan4phi)
  Dim IV As Double: IV = nu * cosphi
  Dim v As Double: v = (nu / 6) * cos3phi * (nu / rho - tan2phi)
  Dim VI As Double: VI = (nu / 120) * cos5phi * (5 - 18 * tan2phi + tan4phi + 14 * eta2 - 58 * tan2phi * eta2)
  
  Dim deltalambda As Double: deltalambda = lambda - lambda0
  Dim deltalambda2 As Double: deltalambda2 = deltalambda * deltalambda
  Dim deltalambda3 As Double: deltalambda3 = deltalambda2 * deltalambda
  Dim deltalambda4 As Double: deltalambda4 = deltalambda3 * deltalambda
  Dim deltalambda5 As Double: deltalambda5 = deltalambda4 * deltalambda
  Dim deltalambda6 As Double: deltalambda6 = deltalambda5 * deltalambda
  
  Dim Northings As Double: Northings = i + II * deltalambda2 + III * deltalambda4 + IIIA * deltalambda6
  Dim Eastings As Double: Eastings = E0 + IV * deltalambda + v * deltalambda3 + VI * deltalambda5
  getEastingsNorthings = Array(Eastings, Northings)
End Function

Private Sub initLatLng()
  ' Airy 1830 major & minor semi-axes
  Const a As Double = 6377563.396, b As Double = 6356256.909
  
  'NatGrid scale factor on central meridian
  Const F0 As Double = 0.9996012717
  
  'NatGrid true origin is 49°N,2°W
  'Both 49 and -2 are converted to radians and hardcoded below
  Const phi0 As Double = 0.85521133348, lambda0 As Double = -0.03490658504
  
  'northing & easting of true origin, metres
  Const N0 As Double = -100000, E0 As Double = 400000
  
  'eccentricity squared
  Dim e2 As Double: e2 = 1 - (b * b) / (a * a)
  
  'n, n², n³
  Dim n As Double: n = (a - b) / (a + b)
  Dim n2 As Double: n2 = n * n
  Dim n3 As Double: n3 = n2 * n
  
  'Loop to find phi and M
  Dim phi As Double: phi = phi0
  Dim M As Double: M = 0
  Do
    phi = (pNorthings - N0 - M) / (a * F0) + phi
    Dim Ma As Double: Ma = (1 + n + (5 / 4) * n2 + (5 / 4) * n3) * (phi - phi0)
    Dim Mb As Double: Mb = (3 * n + 3 * n * n + (21 / 8) * n3) * Sin(phi - phi0) * Cos(phi + phi0)
    Dim Mc As Double: Mc = ((15 / 8) * n2 + (15 / 8) * n3) * Sin(2 * (phi - phi0)) * Cos(2 * (phi + phi0))
    Dim Md As Double: Md = (35 / 24) * n3 * Sin(3 * (phi - phi0)) * Cos(3 * (phi + phi0))
    M = b * F0 * (Ma - Mb + Mc - Md) 'meridional arc
  Loop While Abs(pNorthings - N0 - M) >= 0.00001 'i.e. until < 0.01mm
  
  Dim cosphi As Double: cosphi = Cos(phi)
  Dim sinphi As Double: sinphi = Sin(phi)
  Dim nu As Double: nu = a * F0 / Sqr(1 - e2 * sinphi * sinphi)                   'nu = transverse radius of curvature
  Dim rho As Double: rho = a * F0 * (1 - e2) / ((1 - e2 * sinphi * sinphi) ^ 1.5) ' rho = meridional radius of curvature
  Dim eta2 As Double: eta2 = nu / rho - 1                                         ' eta = ?
  
  Dim tanphi As Double: tanphi = Math.Tan(phi)
  Dim tan2phi As Double: tan2phi = tanphi * tanphi
  Dim tan4phi As Double: tan4phi = tan2phi * tan2phi
  Dim tan6phi As Double: tan6phi = tan4phi * tan2phi
  Dim secphi As Double: secphi = 1 / cosphi
  Dim nu3 As Double: nu3 = nu * nu * nu
  Dim nu5 As Double: nu5 = nu3 * nu * nu
  Dim nu7 As Double: nu7 = nu5 * nu * nu
  Dim VII As Double: VII = tanphi / (2 * rho * nu)
  Dim VIII As Double: VIII = tanphi / (24 * rho * nu3) * (5 + 3 * tan2phi + eta2 - 9 * tan2phi * eta2)
  Dim IX As Double: IX = tanphi / (720 * rho * nu5) * (61 + 90 * tan2phi + 45 * tan4phi)
  Dim X As Double: X = secphi / nu
  Dim XI As Double: XI = secphi / (6 * nu3) * (nu / rho + 2 * tan2phi)
  Dim XII As Double: XII = secphi / (120 * nu5) * (5 + 28 * tan2phi + 24 * tan4phi)
  Dim XIIA As Double: XIIA = secphi / (5040 * nu7) * (61 + 662 * tan2phi + 1320 * tan4phi + 720 * tan6phi)
  
  Dim dE As Double: dE = (pEastings - E0)
  Dim dE2 As Double: dE2 = dE * dE
  Dim dE3 As Double: dE3 = dE2 * dE
  Dim dE4 As Double: dE4 = dE2 * dE2
  Dim dE5 As Double: dE5 = dE3 * dE2
  Dim dE6 As Double: dE6 = dE4 * dE2
  Dim dE7 As Double: dE7 = dE5 * dE2
  phi = phi - VII * dE2 + VIII * dE4 - IX * dE6
  Dim lambda As Double: lambda = lambda0 + X * dE - XI * dE3 + XII * dE5 - XIIA * dE7
  
  'Set latitude and longitude
  pLatLngAv = True
  pLatitude = toDeg(phi)
  pLongitude = toDeg(lambda)
End Sub

Private Function toRad(ByVal deg As Double) As Double
  Const PI As Double = 3.1415926535
  toRad = deg * PI / 180
End Function
Private Function toDeg(ByVal rad As Double) As Double
  Const PI As Double = 3.1415926535
  toDeg = rad * 180 / PI
End Function

Private Function regex(ByVal sPattern As String, Optional ByVal sFlags As String = "") As Object
  Dim rx As Object: Set rx = CreateObject("VbScript.Regexp")
  rx.Pattern = sPattern
  If sFlags <> "" Then
    If InStr(1, sFlags, "i") > 0 Then rx.IgnoreCase = True
    If InStr(1, sFlags, "g") > 0 Then rx.Global = True
  End If
  Set regex = rx
End Function
Private Function floor(ByVal v As Double)
  floor = Int(v) - 1 * (Int(v) > v)
End Function

Private Function gisIPoint_Latitude() As Double
  gisIPoint_Latitude = Latitude
End Function

Private Function gisIPoint_Longitude() As Double
  gisIPoint_Longitude = Longitude
End Function
