import os
import sys
import comtypes.client

class Etabswrapper():
    def __init__(self,ProgramPath='C:\\Program Files\\Computers and Structures\\ETABS 20\\ETABS.exe',APIPath='C:\CSiEtabsAPIexample',ModelName='Example01',unit=6,ActiveObject=True):
        self.ProgramPath  = ProgramPath
        self.APIPath      = APIPath
        self.ModelName    = ModelName
        self.unit         = unit
        self.ActiveObject = ActiveObject
        
        self.Etabsmodel = self.initilaze()
        
        pass


    def initilaze(self):
        """
        ProgramPath='C:\Program Files\Computers and Structures\SAP2000 22\SAP2000.exe',
        APIPath  : 'C:\CSiAPIexample',
        ModelName='Example01'
        unit : switch to k-ft units default kN_m_C

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
        """
        #set the following flag to True to attach to an existing instance of the program
        #otherwise a new instance of the program will be started
        AttachToInstance = True

        #set the following flag to True to manually specify the path to ETABS.exe
        #this allows for a connection to a version of ETABS other than the latest installation
        #otherwise the latest installed version of ETABS will be launched
        SpecifyPath = False

        
        if not os.path.exists(self.APIPath):
            try:
                os.makedirs(self.APIPath)
            except OSError:
                pass
        ModelPath = self.APIPath + os.sep + self.ModelName
        
        #create API helper object
        helper = comtypes.client.CreateObject('ETABSv1.Helper')
        helper = helper.QueryInterface(comtypes.gen.ETABSv1.cHelper)

        if AttachToInstance:
            #attach to a running instance of ETABS
            try:
                #get the active ETABS object
                myETABSObject = helper.GetObject("CSI.ETABS.API.ETABSObject") 
            except (OSError, comtypes.COMError):
                print("No running instance of the program found or failed to attach.")
                sys.exit(-1)
        else:
            if SpecifyPath:
                try:
                    #'create an instance of the ETABS object from the specified path
                    myETABSObject = helper.CreateObject(self.ProgramPath)
                except (OSError, comtypes.COMError):
                    print("Cannot start a new instance of the program from " + self.ProgramPath)
                    sys.exit(-1)
            else:
                try: 
                    #create an instance of the ETABS object from the latest installed ETABS
                    myETABSObject = helper.CreateObjectProgID("CSI.ETABS.API.ETABSObject") 
                except (OSError, comtypes.COMError):
                    print("Cannot start a new instance of the program.")
                    sys.exit(-1)

            
        """
        try:
            myEtabsObject = comtypes.client.GetActiveObject('CSI.ETABS.API.ETABSObject')
        except (OSError,comtypes.COMError):
            self.ActiveObject = False
            print('Yeni bir proje başlatılıyor...')
            myEtabsObject = comtypes.client.CreateObject('ETABSv1.Helper').QueryInterface(comtypes.gen.ETABSv1.cHelper).\
                CreateObjectProgID('CSI.ETABS.API.ETABSObject')
        finally:
            #start ETABS application
            myETABSObject.ApplicationStart()
            EtabsModel = myEtabsObject.SapModel
            if self.ActiveObject == False:
                EtabsModel.InitializeNewModel()
                """
                
        #start ETABS application
        myETABSObject.ApplicationStart()

        #create SapModel object
        SapModel = myETABSObject.SapModel

        #initialize model
        SapModel.InitializeNewModel()
        #create new blank model
        ret = SapModel.File.NewBlank()
        return ret
        
                
        #EtabsModel.SetPresentUnits(self.unit)

        return EtabsModel

    def portalFrame(self,NumberStorys,StoryHeight,NumberBays,BayWidth):
        try:
            self.Etabsmodel.File.New2DFrame(type,NumberStorys,StoryHeight,NumberBays,BayWidth,True)
        except:
            self.exitapp()
            raise("işlem gerçekleştirilemedi")
            
        
    def CreateGridSystem(self):
        self.Etabsmodel.SetPresentUnits(6)
        self.Etabsmodel.File.NewGridOnly(4, 3, 3, 4, 3, 3, 4)
        # self.SapModel.File.NewGridOnly(numberStory, typicalStoryHeight, BottomStoryHeight, numberLinesX, numberLinesY, SpacingX, SpacingY)

    def matNames(self):
        matName = self.SapModel.PropMaterial.GetNameList()
        print(matName)

    def getMatProp(self):
        
        getProp = self.Etabsmodel.PropMaterial.GetOConcrete('C25/30')
        # GetOConcrete(Name, Fc, IsLightWeight, FcsFactor, SSType, SSHysType, StrainAtFc, StrainUltimate, FrictionAngle, DilatationAngle, temp)
        print(getProp)
        getWeight = self.Etabsmodel.PropMaterial.GetWeightAndMass('C25/30')
        print(getWeight)
        getMechProp = self.Etabsmodel.PropMaterial.GetMPIsotropic('C25/30')
        print(getMechProp)

    def createNewMaterial(self):
        self.Etabsmodel.PropMaterial.SetMaterial('C30', 2)
        # SetOConcrete(Name, Fc, IsLightWeight, FcsFactor, SSType, SSHysType, StrainAtFc, StrainUltimate, FrictionAngle, DilatationAngle, temp)
        # Hognestad Method ==> StrainUltimate = 0.0038 ; Etabs StrainUltimate = 0.005
        # StrainAtFc = 2*Fc/E ==> 2 * 30 / 32000 = 0.001875
        self.Etabsmodel.PropMaterial.SetOConcrete('C30', 30000, False, 0, 2, 4, 0.001875, 0.0038, 0, 0, 0)
        self.Etabsmodel.PropMaterial.SetWeightAndMass('C30', 1, 25)
        self.Etabsmodel.PropMaterial.SetMPIsotropic('C30', 32000000, 0.2, 1e-5)

    def deleteMaterial(self):
        self.SapModel.PropMaterial.Delete('C25/30')

    def exitapp(self):
        #close the program
        self.Etabsmodel.ApplicationExit(False)
        sapModel = None
        sapModel = None
        

    
class Sapwrapper():
    def __init__(self):
        pass
        
        

    def initilaze(ProgramPath='C:\Program Files\Computers and Structures\SAP2000 22\SAP2000.exe',APIPath='C:\CSiAPIexample',ModelName='Example01',unit=6):
        """
        ProgramPath='C:\Program Files\Computers and Structures\SAP2000 22\SAP2000.exe',
        APIPath  : 'C:\CSiAPIexample',
        ModelName='Example01'
        unit : switch to k-ft units default kN_m_C

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
        """
        AttachToInstance = False
        SpecifyPath = False

        if not os.path.exists(APIPath):
            try:
                os.makedirs(APIPath)
            except OSError:
                pass

        ModelPath = APIPath + os.sep + ModelName+'.sdb'
        #create API helper object

        helper = comtypes.client.CreateObject('SAP2000v1.Helper')

        helper = helper.QueryInterface(comtypes.gen.SAP2000v1.cHelper)

        if AttachToInstance:
            #attach to a running instance of SAP2000
            try:
                #get the active SapObject
                    mySapObject = helper.GetObject("CSI.SAP2000.API.SapObject") 

            except (OSError, comtypes.client.COMError):
                print("No running instance of the program found or failed to attach.")
                sys.exit(-1)

        else:
            if SpecifyPath:
                try:
                    #'create an instance of the SAPObject from the specified path
                    mySapObject = helper.CreateObject(ProgramPath)

                except (OSError, comtypes.client.COMError):
                    print("Cannot start a new instance of the program from " + ProgramPath)
                    sys.exit(-1)

            else:
                try:
                    #create an instance of the SAPObject from the latest installed SAP2000
                    mySapObject = helper.CreateObjectProgID("CSI.SAP2000.API.SapObject")

                except (OSError, comtypes.client.COMError):
                    print("Cannot start a new instance of the program.")
                    sys.exit(-1)
                    
            #start SAP2000 application

            mySapObject.ApplicationStart()

        #create SapModel object

        SapModel = mySapObject.SapModel

        #initialize model

        SapModel.InitializeNewModel()

        #ret = SapModel.File.OpenFile(FileName) mevcut dosyayı açmak için

        SapModel.SetPresentUnits(unit)

        return SapModel

    def portalFrame(sapModel,type,NumberStorys,StoryHeight,NumberBays,BayWidth):
        """
        Parameters
            type : One of the following 2D frame template types in the e2DFrameType enumeration.

                        PortalFrame = 0

                        ConcentricBraced = 1

                        EccentricBraced = 2

            NumberStorys : The number of stories in the frame.

            StoryHeight : The height of each story. [L]

            NumberBays : The number of bays in the frame.

            BayWidth : The width of each bay. [L]


            Restraint : Joint restraints are provided at the base of the frame when this item is True.

            Beam : The frame section property used for all beams in the frame. This must either be Default or the name of a defined frame section property.

            Column : The frame section property used for all columns in the frame. This must either be Default or the name of a defined frame section property.

            Brace : The frame section property used for all braces in the frame. This must either be Default or the name of a defined frame section property.
                     This item does not apply to the portal frame.

        """
        sapModel.File.New2DFrame(type,NumberStorys,StoryHeight,NumberBays,BayWidth,True)

    def create_material(sapModel,material,materialName = 'CONC',E=3600,u=0.2,thermal_coef=0.0000055):
        """
        material 1-7 arasında değerler alabilir aşağıda değerlere karşılık gelen malzemeler belirtilmiştir.
        #define material property

                eMatType_Steel = 1

                eMatType_Concrete = 2

                eMatType_NoDesign = 3

                eMatType_Aluminum = 4

                eMatType_ColdFormed = 5

                eMatType_Rebar = 6

                eMatType_Tendon = 7
        E :Elastisite module
        u :Poisson’s ratio.
        thermal_coef :  thermal coefficient
            This item applies only if the specified material has properties that are temperature dependent. That is, it applies only if properties are specified for the material at more than one temperature.

            This item is the temperature at which the specified data applies. The temperature must have been defined previously for the material.

        """
        sapModel.PropMaterial.SetMaterial(materialName, material)

        sapModel.PropMaterial.SetMPIsotropic(materialName,E,u,thermal_coef)

        """
        Detaylı malzeme tanımlama yöntemleri bunlara detaylı bakılması lazım aşağıda örnek bir tane bulunuyor
        Function SetOConcrete_1(ByVal Name As String, ByVal fc As Double, ByVal IsLightweight As Boolean, ByVal fcsfactor As Double, ByVal sstype As Long, ByVal SSHysType As Long, ByVal StrainAtfc As Double, ByVal StrainUltimate As Double, ByVal FinalSlope As Double, Optional ByVal FrictionAngle As Double = 0, Optional ByVal DilatationalAngle As Double = 0, Optional ByVal Temp As Double = 0) As Long

        Parameters
        Name

        The name of an existing concrete material property.

        fc

        The concrete compressive strength. [F/L2]

        IsLightweight

        If this item is True, the concrete is assumed to be lightweight concrete.

        fcsfactor

        The shear strength reduction factor for lightweight concrete.

        SSType

        This is 0, 1 or 2, indicating the stress-strain curve type.

                0 = User defined

                1 = Parametric - Simple

                2 = Parametric - Mander

        SSHysType

        This is 0, 1 or 2, indicating the stress-strain hysteresis type.

                0 = Elastic

                1 = Kinematic

                2 = Takeda

        StrainAtfc

        This item applies only to parametric stress-strain curves. It is the strain at the unconfined compressive strength.

        StrainUltimate

        This item applies only to parametric stress-strain curves. It is the ultimate unconfined strain capacity. This item must be larger than the StrainAtfc item.

        FinalSlope

        This item applies only to parametric stress-strain curves. It is a multiplier on the material modulus of elasticity, E. This value multiplied times E gives the final slope on the compression side of the curve.

        FrictionAngle

        The Drucker-Prager friction angle, 0 <= FrictionAngle < 90. [deg]

        DilatationalAngle

        The Drucker-Prager dilatational angle, 0 <= DilatationalAngle < 90. [deg]

        Temp

        This item applies only if the specified material has properties that are temperature dependent. That is, it applies only if properties are specified for the material at more than one temperature.

        This item is the temperature at which the specified data applies. The temperature must have been defined previously for the material.


        """

    def rectang(sapModel,sectionName,materialName,t3,t2,modValue= [1000, 0, 0, 1, 1, 1, 1, 1]):
        """
        VB6 Procedure
            Function SetRectangle(Name As String,MatProp As String,t3 As Double,t2 As Double, Optional ByVal Color As Long = -1, Optional ByVal Notes As String = "", Optional ByVal GUID As String = "") As Long

            Parameters
            Name

            The name of an existing or new frame section property. If this is an existing property, that property is modified; otherwise, a new property is added.

            MatProp

            The name of the material property for the section.

            t3

            The section depth. [L]

            t2

            The section width. [L]

            modValue = list ["Cross-section(axial)Area","Shear Area in 2 direction","Shear Area in 3 direction",
        "Torsional Constant","Moment of Inertia about 2 axis","Moment of Inertia about 3 axis",
        "Mass","Weight"
        ]
                example => modValue = [1000, 0, 0, 1, 1, 1, 1, 1]

            Color

            The display color assigned to the section. If Color is specified as -1, the program will automatically assign a color.

            Notes

            The notes, if any, assigned to the section.

            GUID

            The GUID (global unique identifier), if any, assigned to the section. If this item is input as Default, the program assigns a GUID to the section.


        """
        sapModel.PropFrame.SetRectangle(sectionName, materialName,t3,t2)
        #define frame section property modifiers        

        sapModel.PropFrame.SetModifiers(sectionName, modValue)

        #'add ASTM A706 rebar material
        #ret = SapModel.PropMaterial.AddQuick(RebarName, MATERIAL_REBAR, , , , , MATERIAL_REBAR_SUBTYPE_ASTM_A706)
        #'set column rebar data
        #ret = SapModel.PropFrame.SetRebarColumn("R1", RebarName, RebarName, 2, 2, 2, 10, 0, 0, "#10", "#5", 4, 0, 0, False)
        
    def addLoadPattern(sapModel,patternName,patternType,selfMultiplier=0,addLoadCase=True):
        """
        Function Add(Name String,MyType eLoadPatternType,Optional SelfWTMultiplier Double = 0,
         Optional AddLoadCase Boolean = True) As Long

            Parameters
            Name

            The name for the new load pattern.

            MyType

            This is one of the following items in the eLoadPatternType enumeration:

            LTYPE_DEAD = 1
            LTYPE_SUPERDEAD = 2
            LTYPE_LIVE = 3
            LTYPE_REDUCELIVE = 4
            LTYPE_QUAKE = 5
            LTYPE_WIND= 6
            LTYPE_SNOW = 7
            LTYPE_OTHER = 8
            LTYPE_MOVE = 9
            LTYPE_TEMPERATURE = 10
            LTYPE_ROOFLIVE = 11
            LTYPE_NOTIONAL = 12
            LTYPE_PATTERNLIVE = 13
            LTYPE_WAVE= 14
            LTYPE_BRAKING = 15
            LTYPE_CENTRIFUGAL = 16
            LTYPE_FRICTION = 17
            LTYPE_ICE = 18
            LTYPE_WINDONLIVELOAD = 19
            LTYPE_HORIZONTALEARTHPRESSURE = 20
            LTYPE_VERTICALEARTHPRESSURE = 21
            LTYPE_EARTHSURCHARGE = 22
            LTYPE_DOWNDRAG = 23
            LTYPE_VEHICLECOLLISION = 24
            LTYPE_VESSELCOLLISION = 25
            LTYPE_TEMPERATUREGRADIENT = 26
            LTYPE_SETTLEMENT = 27
            LTYPE_SHRINKAGE = 28
            LTYPE_CREEP = 29
            LTYPE_WATERLOADPRESSURE = 30
            LTYPE_LIVELOADSURCHARGE = 31
            LTYPE_LOCKEDINFORCES = 32
            LTYPE_PEDESTRIANLL = 33
            LTYPE_PRESTRESS = 34
            LTYPE_HYPERSTATIC = 35
            LTYPE_BOUYANCY = 36
            LTYPE_STREAMFLOW = 37
            LTYPE_IMPACT = 38
            LTYPE_CONSTRUCTION = 39
            
            SelfWTMultiplier
            The self weight multiplier for the new load pattern.

            AddLoadCase
            If this item is True, a linear static load case corresponding to the new load pattern is added.
        """
        sapModel.LoadPatterns.Add(patternName, patternType,selfMultiplier,addLoadCase)
    
    def save(sapModel,APIPath='C:\CSiAPIexample',ModelName='Example01'):
        #save model
        ModelPath = APIPath+ModelName
        sapModel.File.Save(ModelPath)
    
    def runAnalysis(sapModel):
        #run model (this will create the analysis model)
        sapModel.Analyze.RunAnalysis()
    
    def exitapp(sapModel):
        #close the program
        sapModel.ApplicationExit(False)
        sapModel = None
        sapModel = None