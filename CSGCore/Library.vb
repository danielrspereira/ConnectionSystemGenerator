Public Class Library
    Public Enum TemplateParts
        BOW1 = 774198
        BOW9 = 808178
        ORIBITALBOW1 = 808394
        STUTZEN4 = 774199
        STUTZEN5 = 774200
        STUTZEN8 = 774201
        STUTZEN45IN = 826307
        STUTZEN45OUT = 807255
    End Enum

    Public Enum JobStatus
        INQUEUE = 0
        INWORK = 1
        FINISHED = 100
        FAILED = -100
        NOTRANSFER = -99
        NOSAVECMD = -98
        NOADDIN = -97
        NOENV = -96
        ERRORSAVE = -1
        ERRORCIRC = -2
    End Enum
End Class

Public Class UnitData
    Public Property ApplicationType As String = ""
    Public Property BOMList As New List(Of BOMItem)
    Public Property Coillist As New List(Of CoilData)
    Public Property DecSeperator As String = ""
    Public Property HasBrineDefrost As Boolean
    Public Property HasIntegratedSubcooler As Boolean
    Public Property IsProdavit As Boolean
    Public Property ModelRangeName As String = ""
    Public Property ModelRangeSuffix As String = ""
    Public Property MultiCircuitDesign As String = ""
    Public Property Occlist As New List(Of PartData)
    Public Property OrderData As New JobInfo
    Public Property CopiedFrom As New JobInfo
    Public Property PEDCategory As Integer
    Public Property SELanguageID As Integer
    Public Property TubeSheet As New SheetData
    Public Property CoverSheet As New SheetData
    Public Property UnitDescription As String = ""
    Public Property UnitSize As String = ""
    Public Property UnitFile As New FileData
    Public Property DLLVersion As String = ""
    Public Property APPVersion As String = ""
End Class

Public Class CoilData
    Public Property Alignment As String = ""
    Public Property Backbowids As New List(Of String)
    Public Property BOMItem As New BOMItem
    Public Property Circuits As New List(Of CircuitData)
    Public Property CoilFile As New FileData
    Public Property ConSyss As New List(Of ConSysData)
    Public Property Defrostbows As List(Of Double)()
    Public Property EDefrostPDMID As String = ""
    Public Property FinnedDepth As Double
    Public Property FinnedHeight As Double
    Public Property FinnedLength As Double
    Public Property Frontbowids As New List(Of String)
    Public Property Gap As Double
    Public Property NoBlindTubes As Integer
    Public Property NoBlindTubeLayers As Integer
    Public Property NoLayers As Integer
    Public Property NoRows As Integer
    Public Property Number As Integer
    Public Property Occlist As New List(Of PartData)
    Public Property RotationDirection As Integer
    Public Property SupportTubesPosition As List(Of Double)()
End Class

Public Class CircuitData
    Public Property Backlevel As Integer
    Public Property Backbowids As New List(Of String)
    Public Property CircuitFile As New FileData
    Public Property CircuitNumber As Integer
    Public Property CircuitSize As Double()
    Public Property CircuitType As String = ""
    Public Property Coilnumber As Integer
    Public Property ConnectionSide As String = ""
    Public Property CoreTube As TubeData
    Public Property CoreTubes As New List(Of CoreTubeLocations)
    Public Property CoreTubeOverhang As Double
    Public Property CustomCirc As Boolean = False
    Public Property FinType As String = ""
    Public Property FinName As String = ""
    Public Property FrontLevel As Integer
    Public Property Frontbowids As New List(Of String)
    Public Property IsOnebranchEvap As Boolean
    Public Property Hairpins As New List(Of HairpinData)
    Public Property L1Levels As New List(Of Double)
    Public Property NoDistributions As Integer
    Public Property NoPasses As Integer
    Public Property Occlist As New List(Of PartData)
    Public Property Orbitalwelding As Boolean = False
    Public Property PEDCategory As Integer
    Public Property PitchX As Double
    Public Property PitchY As Double
    Public Property PDMID As String = ""
    Public Property Pressure As Integer
    Public Property Quantity As Integer
    Public Property SupportPDMID As String = ""
End Class

Public Class ConSysData
    Public Property BOMItem As New BOMItem
    Public Property Circnumber As Integer
    Public Property ConSysFile As New FileData
    Public Property ControlNipples As Boolean = True
    Public Property ConType As Integer          'threads vs flanges 
    Public Property CoreTubes As New List(Of CoreTubeLocations)
    Public Property CoverSheetCutouts As New List(Of NippleCutoutData)
    Public Property FlangeDims As New FlangeData
    Public Property FlangeID As String = ""
    Public Property HasFTCon As Boolean
    Public Property HasHotgas As Boolean
    Public Property HasBallValve As Boolean
    Public Property HasSensor As Boolean
    Public Property HotGasData As New HotgasInfo
    Public Property HotGasConnectionDiameter As Double
    Public Property HeaderAlignment As String = ""
    Public Property HeaderMaterial As String = ""
    Public Property InletConjunctions As New List(Of TubeData)
    Public Property InletHeaders As New List(Of HeaderData)
    Public Property InletNipples As New List(Of TubeData)
    Public Property Occlist As New List(Of PartData)
    Public Property OilSifons As New List(Of HeaderData)
    Public Property OutletConjunctions As New List(Of TubeData)
    Public Property OutletHeaders As New List(Of HeaderData)
    Public Property OutletNipples As New List(Of TubeData)
    Public Property SpecialCX As Boolean = False
    Public Property SpecialRX As Boolean = False
    Public Property Valvesize As String = ""
    Public Property VType As String = ""
End Class

Public Class TubeData
    Public Property Angle As Double
    Public Property BottomCapNeeded As Boolean = False
    Public Property BottomCapID As String = ""
    Public Property TopCapNeeded As Boolean = False
    Public Property TopCapID As String = ""
    Public Property Diameter As Double
    Public Property FileName As String = ""
    Public Property HeaderType As String = ""
    Public Property IsBrine As Boolean
    Public Property Material As String = ""
    Public Property Materialcodeletter As String = ""
    Public Property Length As Double
    Public Property Quantity As Integer
    Public Property RawLength As Double
    Public Property RawMaterial As String = ""
    Public Property TubeFile As New FileData
    Public Property TubeType As String = ""     'header, nipple
    Public Property SVPosition As String() = {"", ""}
    Public Property WallThickness As Double
End Class

Public Class HeaderData
    Public Property Dim_a As Double
    Public Property Displacehor As Double
    Public Property Displacever As Double
    Public Property Nippletubes As Integer
    Public Property Nipplepositions As New List(Of Double)
    Public Property OddLocation As String = ""
    Public Property Origin As Double()
    Public Property Overhangtop As Double
    Public Property Overhangbottom As Double
    Public Property StutzenDatalist As New List(Of StutzenData)
    Public Property Tube As New TubeData
    Public Property VentIDs As String()
    Public Property Ventposition As Double
    Public Property Ventsize As String = ""
    Public Property Xlist As New List(Of Double)
    Public Property Ylist As New List(Of Double)
End Class

Public Class StutzenData
    Public Property ID As String = ""
    Public Property Angle As Double
    Public Property ABV As Double
    Public Property SpecialTag As String = ""
    Public Property HoleOffset As Double
    Public Property Figure As Integer
    Public Property XPos As Double
    Public Property YPos As Double
    Public Property ZPos As Double
End Class

Public Class SheetData
    Public Property Dim_d As Double
    Public Property IsPowderCoated As Boolean
    Public Property Material As String = ""
    Public Property MaterialCodeLetter As String = ""
    Public Property PunchingType As Integer
    Public Property Thickness As Double
End Class

Public Class JobInfo
    Public Property Uid As Integer
    Public Property OrderNumber As String
    Public Property OrderPosition As String
    Public Property ProjectNumber As String = ""
    Public Property Prio As Integer
    Public Property Status As Integer
    Public Property Path As String = ""
    Public Property Plant As String = ""
    Public Property Users As String = ""
    Public Property RequestTime As Date
    Public Property FinishedTime As Date
    Public Property ModelRange As String = ""
    Public Property PDMID As String = ""
    Public Property Workspace As String = ""
    Public Property OrderDir As String = ""
    Public Property IsERPTest As Boolean = False
    Public Property IsPDMTest As Boolean = False
End Class

Public Class FileData
    Public Property AGPno As String = ""
    Public Property CDB_de As String = ""
    Public Property CDB_en As String = ""
    Public Property CDB_Material As String = ""
    Public Property CDB_z_Bemerkung As String = ""
    Public Property CDB_Zusatzbenennung As String = ""
    Public Property CSG As String = "1"
    Public Property LNCode As String = ""
    Public Property Filetype As String = ""
    Public Property Fullfilename As String = ""
    Public Property Orderno As String = ""
    Public Property Orderpos As String = ""
    Public Property Plant As String = ""
    Public Property Projectno As String = ""
    Public Property Shortname As String = ""
    Public Property Z_Kategorie As String = ""
End Class

Public Class BowData
    Public Property Ps As Double()
    Public Property Pe As Double()
    Public Property PLCool As New List(Of Integer)
    Public Property PLHot As New List(Of Integer)
    Public Property ID As String
    Public Property L1 As Double
    Public Property Length As Double
    Public Property Wallthickness As Double
    Public Property Level As Integer
    Public Property Uniquekey As String
End Class

Public Class HairpinData
    Public Property PDMID As String = ""
    Public Property ERPCode As String = ""
    Public Property Pitch As Double
    Public Property RefBow As String = ""
End Class

Public Class PartData
    Public Property Occindex As Integer
    Public Property Occname As String = ""
    Public Property Configref As String
    Public Property BOMAGP As String = ""
End Class

Public Class CoreTubeLocations
    Public Property Shortname As String = ""
    Public Property Xlist As New List(Of Double)
    Public Property Ylist As New List(Of Double)
    Public Property Zlist As New List(Of Double)
    Public Property Xindizes As New List(Of List(Of Integer))
    Public Property Yindizes As New List(Of List(Of Integer))
    Public Property Zindizes As New List(Of List(Of Integer))
End Class

Public Class PlateData
    Public Property ID As String = ""
    Public Property Quantity As Integer
    Public Property InnerDiameter As Double
End Class

Public Class BOMItem
    Public Property Item As String = ""
    Public Property ItemDescription As String = ""
    Public Property ItemNumber As Integer
    Public Property ParentItemNumber As Integer
    Public Property Position As Integer
    Public Property Quantity As Double
    Public Property Warehouse As String = ""
    Public Property ItemGroup As String = ""
    Public Property OptionTag As String = ""
End Class

Public Class NippleCutoutData
    Public Property ERPCode As String = ""
    Public Property PDMID As String = ""
    Public Property Filename As String = ""
    Public Property CutsizeY As Double
    Public Property CutsizeZ As Double
    Public Property YPos As Double
    Public Property ZPos As Double
    Public Property Diameter As Double
    Public Property Displacement As Double
    Public Property HasFlange As Boolean = False
    Public Property Alignment As String = ""
    Public Property Parentfile As String = ""
End Class

Public Class DVLineElement
    Public Property DVLine As SolidEdgeDraft.DVLine2d
    Public Property RefFileName As String
End Class

Public Class StutzenDVElement
    Public Property DVElement As Object
    Public Property Xpos As Double
    Public Property Ypos As Double
End Class

Public Class KeyInfo
    Public Property Kindex As Integer
    Public Property X As Double
    Public Property Y As Double
    Public Property Z As Double
    Public Property Ktype As Integer
    Public Property Kkey As String
End Class

Public Class HotgasInfo
    Public Property Diameter As Double
    Public Property Angle As Double
    Public Property Headertype As String = ""
End Class

Public Class FlangeData
    Public Property HF As Double
    Public Property SB As Double
    Public Property Length As Double
    Public Property DF As Double
End Class

Public Class PointData
    Public Property X As Double
    Public Property Y As Double
End Class