import os

import sys

import comtypes.client


#set the following flag to True to attach to an existing instance of the program

#otherwise a new instance of the program will be started

#programın mevcut bir örneğine eklemek için aşağıdaki bayrağı True olarak ayarlayın

#aksi takdirde programın yeni bir örneği başlatılacaktır

AttachToInstance = False

 

#set the following flag to True to manually specify the path to SAP2000.exe

#this allows for a connection to a version of SAP2000 other than the latest installation

#otherwise the latest installed version of SAP2000 will be launched


#SAP2000.exe yolunu manuel olarak belirtmek için aşağıdaki bayrağı True olarak ayarlayın

#bu, en son kurulum dışında bir SAP2000 sürümüne bağlantıya izin verir

#aksi takdirde SAP2000'in en son kurulu sürümü kullanıma sunulacaktır

SpecifyPath = False

 

#if the above flag is set to True, specify the path to SAP2000 below
#yukarıdaki işaret True olarak ayarlanmışsa, aşağıda SAP2000 yolunu belirtin
ProgramPath = 'C:\Program Files\Computers and Structures\SAP2000 22\SAP2000.exe'

 

#full path to the model

#set it to the desired path of your model
#modele giden tam yol

# onu modelinizin istediğiniz yoluna ayarlayın

APIPath = 'C:\CSiAPIexample'

if not os.path.exists(APIPath):

        try:

            os.makedirs(APIPath)

        except OSError:

            pass

ModelPath = APIPath + os.sep + 'API_1-001.sdb'

 

#create API helper object

helper = comtypes.client.CreateObject('SAP2000v1.Helper')

helper = helper.QueryInterface(comtypes.gen.SAP2000v1.cHelper)

 

if AttachToInstance:

    #attach to a running instance of SAP2000

    try:

        #get the active SapObject

            mySapObject = helper.GetObject("CSI.SAP2000.API.SapObject") 

    except (OSError, comtypes.COMError):

        print("No running instance of the program found or failed to attach.")

        sys.exit(-1)

else:

 

    if SpecifyPath:

        try:

            #'create an instance of the SAPObject from the specified path

            mySapObject = helper.CreateObject(ProgramPath)

        except (OSError, comtypes.COMError):

            print("Cannot start a new instance of the program from " + ProgramPath)

            sys.exit(-1)

    else:

        try:

            #create an instance of the SAPObject from the latest installed SAP2000

            mySapObject = helper.CreateObjectProgID("CSI.SAP2000.API.SapObject")

        except (OSError, comtypes.COMError):

            print("Cannot start a new instance of the program.")

            sys.exit(-1)

 

    #start SAP2000 application

    mySapObject.ApplicationStart()

    

#create SapModel object

SapModel = mySapObject.SapModel

 

#initialize model

SapModel.InitializeNewModel()

 

#create new blank model

ret = SapModel.File.NewBlank()

 

#define material property

MATERIAL_CONCRETE = 2

ret = SapModel.PropMaterial.SetMaterial('CONC', MATERIAL_CONCRETE)

 

#assign isotropic mechanical properties to material

ret = SapModel.PropMaterial.SetMPIsotropic('CONC', 3600, 0.2, 0.0000055)

 

#define rectangular frame section property

ret = SapModel.PropFrame.SetRectangle('R1', 'CONC', 12, 12)

 

#define frame section property modifiers

ModValue = [1000, 0, 0, 1, 1, 1, 1, 1]

ret = SapModel.PropFrame.SetModifiers('R1', ModValue)

 

#switch to k-ft units

lb_in_F = 1

lb_ft_F = 2

kip_in_F = 3

kip_ft_F = 4

kN_mm_C = 5

kN_m_C = 6

kgf_mm_C = 7

kgf_m_C = 8

N_mm_C = 9

N_m_C = 10

Ton_mm_C = 11

Ton_m_C = 12

kN_cm_C = 13

kgf_cm_C = 14

N_cm_C = 15

Ton_cm_C = 16


ret = SapModel.SetPresentUnits(kN_m_C)

 

#add frame object by coordinates

FrameName1 = ' '

FrameName2 = ' '

FrameName3 = ' '

[FrameName1, ret] = SapModel.FrameObj.AddByCoord(0, 0, 0, 0, 0, 10, FrameName1, 'R1', '1', 'Global')

[FrameName2, ret] = SapModel.FrameObj.AddByCoord(0, 0, 10, 8, 0, 16, FrameName2, 'R1', '2', 'Global')

[FrameName3, ret] = SapModel.FrameObj.AddByCoord(-4, 0, 10, 0, 0, 10, FrameName3, 'R1', '3', 'Global')

 

#assign point object restraint at base

PointName1 = ' '

PointName2 = ' '

Restraint = [True, True, True, True, False, False]

"""
Parameters
xi, yi, zi

The coordinates of the I-End of the added frame object. The coordinates are in the coordinate system defined by the CSys item.

xj, yj, zj

The coordinates of the J-End of the added frame object. The coordinates are in the coordinate system defined by the CSys item.

Name

This is the name that the program ultimately assigns for the frame object. If no UserName is specified, the program assigns a default name to the frame object. If a UserName is specified and that name is not used for another frame, cable or tendon object, the UserName is assigned to the frame object, otherwise a default name is assigned to the frame object.

PropName

This is Default, None, or the name of a defined frame section property.

If it is Default, the program assigns a default section property to the frame object. If it is None, no section property is assigned to the frame object. If it is the name of a defined frame section property, that property is assigned to the frame object.

UserName

This is an optional user specified name for the frame object. If a UserName is specified and that name is already used for another frame object, the program ignores the UserName.

CSys

The name of the coordinate system in which the frame object end point coordinates are defined.

"""

[PointName1, PointName2, ret] = SapModel.FrameObj.GetPoints(FrameName1, PointName1, PointName2)

ret = SapModel.PointObj.SetRestraint(PointName1, Restraint)

 

#assign point object restraint at top

Restraint = [True, True, False, False, False, False]

[PointName1, PointName2, ret] = SapModel.FrameObj.GetPoints(FrameName2, PointName1, PointName2)

ret = SapModel.PointObj.SetRestraint(PointName2, Restraint)

 

#refresh view, update (initialize) zoom

ret = SapModel.View.RefreshView(0, False)

 

#add load patterns

LTYPE_OTHER = 8

ret = SapModel.LoadPatterns.Add('1', LTYPE_OTHER, 1, True)

ret = SapModel.LoadPatterns.Add('2', LTYPE_OTHER, 0, True)

ret = SapModel.LoadPatterns.Add('3', LTYPE_OTHER, 0, True)

ret = SapModel.LoadPatterns.Add('4', LTYPE_OTHER, 0, True)

ret = SapModel.LoadPatterns.Add('5', LTYPE_OTHER, 0, True)

ret = SapModel.LoadPatterns.Add('6', LTYPE_OTHER, 0, True)

ret = SapModel.LoadPatterns.Add('7', LTYPE_OTHER, 0, True)

 

#assign loading for load pattern 2

[PointName1, PointName2, ret] = SapModel.FrameObj.GetPoints(FrameName3, PointName1, PointName2)

PointLoadValue = [0,0,-10,0,0,0]

ret = SapModel.PointObj.SetLoadForce(PointName1, '2', PointLoadValue)

ret = SapModel.FrameObj.SetLoadDistributed(FrameName3, '2', 1, 10, 0, 1, 1.8, 1.8)

 

#assign loading for load pattern 3

[PointName1, PointName2, ret] = SapModel.FrameObj.GetPoints(FrameName3, PointName1, PointName2)

PointLoadValue = [0,0,-17.2,0,-54.4,0]

ret = SapModel.PointObj.SetLoadForce(PointName2, '3', PointLoadValue)

 

#assign loading for load pattern 4

ret = SapModel.FrameObj.SetLoadDistributed(FrameName2, '4', 1, 11, 0, 1, 2, 2)

 

#assign loading for load pattern 5

ret = SapModel.FrameObj.SetLoadDistributed(FrameName1, '5', 1, 2, 0, 1, 2, 2, 'Local')

ret = SapModel.FrameObj.SetLoadDistributed(FrameName2, '5', 1, 2, 0, 1, -2, -2, 'Local')

 

#assign loading for load pattern 6

ret = SapModel.FrameObj.SetLoadDistributed(FrameName1, '6', 1, 2, 0, 1, 0.9984, 0.3744, 'Local')

ret = SapModel.FrameObj.SetLoadDistributed(FrameName2, '6', 1, 2, 0, 1, -0.3744, 0, 'Local')

 

#assign loading for load pattern 7

ret = SapModel.FrameObj.SetLoadPoint(FrameName2, '7', 1, 2, 0.5, -15, 'Local')

 

#switch to k-in units

kip_in_F = 3

ret = SapModel.SetPresentUnits(kip_in_F)

 

#save model

ret = SapModel.File.Save(ModelPath)

 

#run model (this will create the analysis model)

ret = SapModel.Analyze.RunAnalysis()

 

#initialize for Sap2000 results

SapResult= [0,0,0,0,0,0,0]

[PointName1, PointName2, ret] = SapModel.FrameObj.GetPoints(FrameName2, PointName1, PointName2)

 

#get Sap2000 results for load cases 1 through 7

for i in range(0,7):

      NumberResults = 0

      Obj = []

      Elm = []

      ACase = []

      StepType = []

      StepNum = []

      U1 = []

      U2 = []

      U3 = []

      R1 = []

      R2 = []

      R3 = []

      ObjectElm = 0

      ret = SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput()

      ret = SapModel.Results.Setup.SetCaseSelectedForOutput(str(i+1))

      if i <= 3:

          [NumberResults, Obj, Elm, ACase, StepType, StepNum, U1, U2, U3, R1, R2, R3, ret] = SapModel.Results.JointDispl(PointName2, ObjectElm, NumberResults, Obj, Elm, ACase, StepType, StepNum, U1, U2, U3, R1, R2, R3)

          SapResult[i] = U3[0]

      else:

          [NumberResults, Obj, Elm, ACase, StepType, StepNum, U1, U2, U3, R1, R2, R3, ret] = SapModel.Results.JointDispl(PointName1, ObjectElm, NumberResults, Obj, Elm, ACase, StepType, StepNum, U1, U2, U3, R1, R2, R3)

          SapResult[i] = U1[0]

 

#close Sap2000

ret = mySapObject.ApplicationExit(False)

SapModel = None

mySapObject = None

 

#fill independent results

IndResult= [0,0,0,0,0,0,0]

IndResult[0] = -0.02639

IndResult[1] = 0.06296

IndResult[2] = 0.06296

IndResult[3] = -0.2963

IndResult[4] = 0.3125

IndResult[5] = 0.11556

IndResult[6] = 0.00651

 

#fill percent difference

PercentDiff = [0,0,0,0,0,0,0]

for i in range(0,7):

      PercentDiff[i] = (SapResult[i] / IndResult[i]) - 1

 

#display results

for i in range(0,7):

      print()

      print(SapResult[i])

      print(IndResult[i])

      print(PercentDiff[i])
