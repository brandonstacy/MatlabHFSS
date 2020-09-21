Dim oHfssApp
Dim oDesktop
Dim oProject
Dim oDesign
Dim oEditor
Dim oModule

Set oHfssApp  = CreateObject("AnsoftHfss.HfssScriptInterface")
Set oDesktop = oHfssApp.GetAppDesktop()
oDesktop.RestoreWindow
oDesktop.NewProject
Set oProject = oDesktop.GetActiveProject

oProject.InsertDesign "HFSS", "SBend", "DrivenModal", ""
Set oDesign = oProject.SetActiveDesign("SBend")
Set oEditor = oDesign.SetActiveEditor("3D Modeler")

oEditor.CreateBox _
Array("NAME:BoxParameters", _
"XPosition:=", "-10.000000mm", _
"YPosition:=", "0.000000mm", _
"ZPosition:=", "-0.017000mm", _
"XSize:=", "20.000000mm", _
"YSize:=", "21.240000mm", _
"ZSize:=", "3.000000mm"), _
Array("NAME:Attributes", _
"Name:=", "AirBox", _
"Flags:=", "", _
"Color:=", "(132 132 193)", _
"Transparency:=", 0.75, _
"PartCoordinateSystem:=", "Global", _
"MaterialName:=", "vacuum", _
"SolveInside:=", true)

oEditor.AssignMaterial _
	Array("NAME:Selections", _
		"Selections:=", "AirBox"), _
	Array("NAME:Attributes", _
		"MaterialName:=", "air", _
		"SolveInside:=", true)

oEditor.ChangeProperty _
Array("NAME:AllTabs", _
	Array("NAME:Geometry3DAttributeTab", _
		Array("NAME:PropServers","AirBox"), _
		Array("NAME:ChangedProps", _
			Array("NAME:Transparent", "Value:=",  1.000000) _
			) _
		) _
	)

oEditor.CreateBox _
Array("NAME:BoxParameters", _
"XPosition:=", "-10.000000mm", _
"YPosition:=", "0.000000mm", _
"ZPosition:=", "-0.017000mm", _
"XSize:=", "20.000000mm", _
"YSize:=", "21.240000mm", _
"ZSize:=", "0.017000mm"), _
Array("NAME:Attributes", _
"Name:=", "GND", _
"Flags:=", "", _
"Color:=", "(132 132 193)", _
"Transparency:=", 0.75, _
"PartCoordinateSystem:=", "Global", _
"MaterialName:=", "vacuum", _
"SolveInside:=", true)

oEditor.AssignMaterial _
	Array("NAME:Selections", _
		"Selections:=", "GND"), _
	Array("NAME:Attributes", _
		"MaterialName:=", "copper", _
		"SolveInside:=", false)

oEditor.ChangeProperty _
	Array("NAME:AllTabs", _
		Array("NAME:Geometry3DAttributeTab", _
			Array("NAME:PropServers", "GND"), _
			Array("NAME:ChangedProps",  _
				Array("NAME:Color", "R:=", 250, "G:=", 250, "B:=", 0) _
			) _
		) _
	) 

oEditor.ChangeProperty _
Array("NAME:AllTabs", _
	Array("NAME:Geometry3DAttributeTab", _
		Array("NAME:PropServers","GND"), _
		Array("NAME:ChangedProps", _
			Array("NAME:Transparent", "Value:=",  0.000000) _
			) _
		) _
	)

oEditor.CreateBox _
Array("NAME:BoxParameters", _
"XPosition:=", "-10.000000mm", _
"YPosition:=", "0.000000mm", _
"ZPosition:=", "-0.017000mm", _
"XSize:=", "20.000000mm", _
"YSize:=", "21.240000mm", _
"ZSize:=", "0.168000mm"), _
Array("NAME:Attributes", _
"Name:=", "Substrate", _
"Flags:=", "", _
"Color:=", "(132 132 193)", _
"Transparency:=", 0.75, _
"PartCoordinateSystem:=", "Global", _
"MaterialName:=", "vacuum", _
"SolveInside:=", true)

oEditor.AssignMaterial _
	Array("NAME:Selections", _
		"Selections:=", "Substrate"), _
	Array("NAME:Attributes", _
		"MaterialName:=", "Rogers RO3003 (tm)", _
		"SolveInside:=", true)

oEditor.ChangeProperty _
	Array("NAME:AllTabs", _
		Array("NAME:Geometry3DAttributeTab", _
			Array("NAME:PropServers", "Substrate"), _
			Array("NAME:ChangedProps",  _
				Array("NAME:Color", "R:=", 50, "G:=", 220, "B:=", 250) _
			) _
		) _
	) 

oEditor.ChangeProperty _
Array("NAME:AllTabs", _
	Array("NAME:Geometry3DAttributeTab", _
		Array("NAME:PropServers","Substrate"), _
		Array("NAME:ChangedProps", _
			Array("NAME:Transparent", "Value:=",  0.400000) _
			) _
		) _
	)

oEditor.CreatePolyline _
Array("NAME:PolylineParameters", "IsPolylineCovered:=", true, "IsPolylineClosed:=", true, _
Array("NAME:PolylinePoints", _
Array("NAME:PLPoint", "X:=", "-0.1900mm", "Y:=", "0.0000mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "-0.1900mm", "Y:=", "2.0000mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "-0.1862mm", "Y:=", "2.0872mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "-0.1748mm", "Y:=", "2.1736mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "-0.1559mm", "Y:=", "2.2588mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "-0.1297mm", "Y:=", "2.3420mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "-0.0963mm", "Y:=", "2.4226mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "-0.0560mm", "Y:=", "2.5000mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "-0.0092mm", "Y:=", "2.5736mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "0.0440mm", "Y:=", "2.6428mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "0.1029mm", "Y:=", "2.7071mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "0.1672mm", "Y:=", "2.7660mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "0.2364mm", "Y:=", "2.8192mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "0.3100mm", "Y:=", "2.8660mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "0.3874mm", "Y:=", "2.9063mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "0.4680mm", "Y:=", "2.9397mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "0.5512mm", "Y:=", "2.9659mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "0.6364mm", "Y:=", "2.9848mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "0.7228mm", "Y:=", "2.9962mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "0.8100mm", "Y:=", "3.0000mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "4.8100mm", "Y:=", "3.0000mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "4.9512mm", "Y:=", "3.0062mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "5.0913mm", "Y:=", "3.0246mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "5.2293mm", "Y:=", "3.0552mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "5.3641mm", "Y:=", "3.0977mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "5.4946mm", "Y:=", "3.1518mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "5.6200mm", "Y:=", "3.2170mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "5.7392mm", "Y:=", "3.2930mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "5.8513mm", "Y:=", "3.3790mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "5.9555mm", "Y:=", "3.4745mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "6.0510mm", "Y:=", "3.5787mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "6.1370mm", "Y:=", "3.6908mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "6.2130mm", "Y:=", "3.8100mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "6.2782mm", "Y:=", "3.9354mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "6.3323mm", "Y:=", "4.0659mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "6.3748mm", "Y:=", "4.2007mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "6.4054mm", "Y:=", "4.3387mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "6.4238mm", "Y:=", "4.4788mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "6.4300mm", "Y:=", "4.6200mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "7.4300mm", "Y:=", "8.6200mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "7.4200mm", "Y:=", "8.8483mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "7.3902mm", "Y:=", "9.0750mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "7.3407mm", "Y:=", "9.2981mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "7.2720mm", "Y:=", "9.5161mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "7.1845mm", "Y:=", "9.7273mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "7.0790mm", "Y:=", "9.9300mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "6.9562mm", "Y:=", "10.1228mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "6.8170mm", "Y:=", "10.3041mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "6.6626mm", "Y:=", "10.4726mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "6.4941mm", "Y:=", "10.6270mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "6.3128mm", "Y:=", "10.7662mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "6.1200mm", "Y:=", "10.8890mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "5.9173mm", "Y:=", "10.9945mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "5.7061mm", "Y:=", "11.0820mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "5.4881mm", "Y:=", "11.1507mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "5.2650mm", "Y:=", "11.2002mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "5.0383mm", "Y:=", "11.2300mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "4.8100mm", "Y:=", "11.2400mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "-0.1900mm", "Y:=", "11.2400mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "-0.5386mm", "Y:=", "11.2552mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "-0.8846mm", "Y:=", "11.3008mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "-1.2253mm", "Y:=", "11.3763mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "-1.5581mm", "Y:=", "11.4812mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "-1.8805mm", "Y:=", "11.6148mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "-2.1900mm", "Y:=", "11.7759mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "-2.4843mm", "Y:=", "11.9634mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "-2.7612mm", "Y:=", "12.1758mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "-3.0184mm", "Y:=", "12.4116mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "-3.2542mm", "Y:=", "12.6688mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "-3.4666mm", "Y:=", "12.9457mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "-3.6541mm", "Y:=", "13.2400mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "-3.8152mm", "Y:=", "13.5495mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "-3.9488mm", "Y:=", "13.8719mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "-4.0537mm", "Y:=", "14.2047mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "-4.1292mm", "Y:=", "14.5454mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "-4.1748mm", "Y:=", "14.8914mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "-4.1900mm", "Y:=", "15.2400mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "-4.1900mm", "Y:=", "21.2400mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "-3.8100mm", "Y:=", "21.2400mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "-3.8100mm", "Y:=", "15.2400mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "-3.7962mm", "Y:=", "14.9245mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "-3.7550mm", "Y:=", "14.6114mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "-3.6867mm", "Y:=", "14.3031mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "-3.5917mm", "Y:=", "14.0019mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "-3.4708mm", "Y:=", "13.7101mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "-3.3250mm", "Y:=", "13.4300mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "-3.1553mm", "Y:=", "13.1637mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "-2.9631mm", "Y:=", "12.9131mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "-2.7497mm", "Y:=", "12.6803mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "-2.5169mm", "Y:=", "12.4669mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "-2.2663mm", "Y:=", "12.2747mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "-2.0000mm", "Y:=", "12.1050mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "-1.7199mm", "Y:=", "11.9592mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "-1.4281mm", "Y:=", "11.8383mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "-1.1269mm", "Y:=", "11.7433mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "-0.8186mm", "Y:=", "11.6750mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "-0.5055mm", "Y:=", "11.6338mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "-0.1900mm", "Y:=", "11.6200mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "4.8100mm", "Y:=", "11.6200mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "5.0715mm", "Y:=", "11.6086mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "5.3309mm", "Y:=", "11.5744mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "5.5865mm", "Y:=", "11.5178mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "5.8361mm", "Y:=", "11.4391mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "6.0779mm", "Y:=", "11.3389mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "6.3100mm", "Y:=", "11.2181mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "6.5307mm", "Y:=", "11.0775mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "6.7384mm", "Y:=", "10.9181mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "6.9313mm", "Y:=", "10.7413mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "7.1081mm", "Y:=", "10.5484mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "7.2675mm", "Y:=", "10.3407mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "7.4081mm", "Y:=", "10.1200mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "7.5289mm", "Y:=", "9.8879mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "7.6291mm", "Y:=", "9.6461mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "7.7078mm", "Y:=", "9.3965mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "7.7644mm", "Y:=", "9.1409mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "7.7986mm", "Y:=", "8.8815mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "7.8100mm", "Y:=", "8.6200mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "6.8100mm", "Y:=", "4.6200mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "6.8024mm", "Y:=", "4.4457mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "6.7796mm", "Y:=", "4.2727mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "6.7419mm", "Y:=", "4.1024mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "6.6894mm", "Y:=", "3.9360mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "6.6226mm", "Y:=", "3.7748mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "6.5421mm", "Y:=", "3.6200mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "6.4483mm", "Y:=", "3.4728mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "6.3421mm", "Y:=", "3.3344mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "6.2242mm", "Y:=", "3.2058mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "6.0956mm", "Y:=", "3.0879mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "5.9572mm", "Y:=", "2.9817mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "5.8100mm", "Y:=", "2.8879mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "5.6552mm", "Y:=", "2.8074mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "5.4940mm", "Y:=", "2.7406mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "5.3276mm", "Y:=", "2.6881mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "5.1573mm", "Y:=", "2.6504mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "4.9843mm", "Y:=", "2.6276mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "4.8100mm", "Y:=", "2.6200mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "0.8100mm", "Y:=", "2.6200mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "0.7560mm", "Y:=", "2.6176mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "0.7023mm", "Y:=", "2.6106mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "0.6495mm", "Y:=", "2.5989mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "0.5979mm", "Y:=", "2.5826mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "0.5480mm", "Y:=", "2.5619mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "0.5000mm", "Y:=", "2.5369mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "0.4544mm", "Y:=", "2.5079mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "0.4115mm", "Y:=", "2.4749mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "0.3716mm", "Y:=", "2.4384mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "0.3351mm", "Y:=", "2.3985mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "0.3021mm", "Y:=", "2.3556mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "0.2731mm", "Y:=", "2.3100mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "0.2481mm", "Y:=", "2.2620mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "0.2274mm", "Y:=", "2.2121mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "0.2111mm", "Y:=", "2.1605mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "0.1994mm", "Y:=", "2.1077mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "0.1924mm", "Y:=", "2.0540mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "0.1900mm", "Y:=", "2.0000mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "0.1900mm", "Y:=", "0.0000mm", "Z:=", "0.1680mm"), _
Array("NAME:PLPoint", "X:=", "-0.1900mm", "Y:=", "0.0000mm", "Z:=", "0.1680mm")), _
Array("NAME:PolylineSegments", _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 0, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 1, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 2, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 3, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 4, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 5, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 6, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 7, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 8, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 9, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 10, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 11, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 12, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 13, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 14, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 15, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 16, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 17, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 18, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 19, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 20, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 21, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 22, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 23, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 24, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 25, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 26, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 27, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 28, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 29, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 30, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 31, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 32, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 33, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 34, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 35, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 36, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 37, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 38, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 39, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 40, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 41, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 42, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 43, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 44, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 45, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 46, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 47, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 48, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 49, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 50, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 51, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 52, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 53, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 54, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 55, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 56, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 57, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 58, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 59, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 60, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 61, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 62, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 63, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 64, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 65, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 66, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 67, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 68, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 69, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 70, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 71, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 72, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 73, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 74, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 75, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 76, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 77, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 78, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 79, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 80, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 81, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 82, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 83, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 84, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 85, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 86, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 87, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 88, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 89, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 90, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 91, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 92, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 93, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 94, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 95, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 96, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 97, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 98, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 99, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 100, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 101, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 102, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 103, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 104, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 105, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 106, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 107, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 108, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 109, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 110, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 111, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 112, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 113, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 114, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 115, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 116, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 117, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 118, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 119, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 120, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 121, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 122, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 123, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 124, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 125, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 126, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 127, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 128, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 129, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 130, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 131, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 132, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 133, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 134, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 135, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 136, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 137, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 138, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 139, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 140, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 141, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 142, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 143, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 144, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 145, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 146, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 147, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 148, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 149, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 150, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 151, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 152, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 153, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 154, "NoOfPoints:=", 2), _
Array("NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 155, "NoOfPoints:=", 2))), _
Array("NAME:Attributes", _
"Name:=", "Microstrip", _
"Flags:=", "", _
"Color:=", "(132 132 193)", _
"Transparency:=", 0.75, _
"PartCoordinateSystem:=", "Global", _
"MaterialName:=", "vacuum", _
"SolveInside:=", true)

oEditor.ThickenSheet _
Array("NAME:Selections", _
"Selections:=", "Microstrip", _
"NewPartsModelFlag:=", _
"Model"), _
Array("NAME:SheetThickenParameters", _
"Thickness:=", "0.017000mm", _
"BothSides:=", false)

oEditor.AssignMaterial _
	Array("NAME:Selections", _
		"Selections:=", "Microstrip"), _
	Array("NAME:Attributes", _
		"MaterialName:=", "copper", _
		"SolveInside:=", false)

oEditor.ChangeProperty _
	Array("NAME:AllTabs", _
		Array("NAME:Geometry3DAttributeTab", _
			Array("NAME:PropServers", "Microstrip"), _
			Array("NAME:ChangedProps",  _
				Array("NAME:Color", "R:=", 250, "G:=", 250, "B:=", 0) _
			) _
		) _
	) 

oEditor.ChangeProperty _
Array("NAME:AllTabs", _
	Array("NAME:Geometry3DAttributeTab", _
		Array("NAME:PropServers","Microstrip"), _
		Array("NAME:ChangedProps", _
			Array("NAME:Transparent", "Value:=",  0.000000) _
			) _
		) _
	)

oEditor.CreateRectangle _
Array("NAME:RectangleParameters", _
"IsCovered:=", true, _
"XStart:=", "-1.000000mm", _
"YStart:=", "0.000000mm", _
"ZStart:=", "-0.017000mm", _
"Width:=", "1.000000mm", _
"Height:=", "2.000000mm", _
"WhichAxis:=", "Y"), _
Array("NAME:Attributes", _
"Name:=", "Port1", _
"Flags:=", "", _
"Color:=", "(132 132 193)", _
"Transparency:=", 7.500000e-01, _
"PartCoordinateSystem:=", "Global", _
"MaterialName:=", "vacuum", _
"SolveInside:=", true)

Set oModule = oDesign.GetModule("BoundarySetup") 

oModule.AssignWavePort _
Array( _
"NAME:WavePort1", _
"NumModes:=", 1, _
"PolarizeEField:=",  false, _
"DoDeembed:=", false, _
"DoRenorm:=", false, _
Array("NAME:Modes", _
Array("NAME:Mode1", _
"ModeNum:=",  1, _
"UseIntLine:=", false) _
), _
"Objects:=", Array("Port1")) 

oEditor.CreateRectangle _
Array("NAME:RectangleParameters", _
"IsCovered:=", true, _
"XStart:=", "-5.000000mm", _
"YStart:=", "21.240000mm", _
"ZStart:=", "-0.017000mm", _
"Width:=", "1.000000mm", _
"Height:=", "2.000000mm", _
"WhichAxis:=", "Y"), _
Array("NAME:Attributes", _
"Name:=", "Port2", _
"Flags:=", "", _
"Color:=", "(132 132 193)", _
"Transparency:=", 7.500000e-01, _
"PartCoordinateSystem:=", "Global", _
"MaterialName:=", "vacuum", _
"SolveInside:=", true)

Set oModule = oDesign.GetModule("BoundarySetup") 

oModule.AssignWavePort _
Array( _
"NAME:WavePort2", _
"NumModes:=", 1, _
"PolarizeEField:=",  false, _
"DoDeembed:=", false, _
"DoRenorm:=", false, _
Array("NAME:Modes", _
Array("NAME:Mode1", _
"ModeNum:=",  1, _
"UseIntLine:=", false) _
), _
"Objects:=", Array("Port2")) 
Set oModule = oDesign.GetModule("BoundarySetup")
oModule.AssignRadiation _
Array("NAME:RadBox", _
"Objects:=", Array("AirBox"), _
"UseAdaptiveIE:=", false)

Set oModule = oDesign.GetModule("AnalysisSetup")
oModule.InsertSetup "HfssDriven", _
Array("NAME:Setup1", _
"Frequency:=", "40.000000GHz", _
"PortsOnly:=", false, _
"maxDeltaS:=", 0.020000, _
"UseMatrixConv:=", false, _
"MaximumPasses:=", 25, _
"MinimumPasses:=", 1, _
"MinimumConvergedPasses:=", 1, _
"PercentRefinement:=", 20, _
"ReducedSolutionBasis:=", false, _
"DoLambdaRefine:=", true, _
"DoMaterialLambda:=", true, _
"Target:=", 0.3333, _
"PortAccuracy:=", 2, _
"SetPortMinMaxTri:=", false)

Set oModule = oDesign.GetModule("AnalysisSetup")
oModule.InsertDrivenSweep _
"Setup1", _
Array("NAME:Sweep1", _
"Type:=", "Interpolating", _
"InterpTolerance:=", 0.500000, _
"InterpMaxSolns:=", 101, _
"SetupType:=", "LinearCount", _
"StartFreq:=", "20.000000GHz", _
"StopFreq:=", "60.000000GHz", _
"Count:=", 120, _
"SaveFields:=", false, _
"ExtrapToDC:=", false)

oProject.SaveAs _
    "C:\Users\bstac\Desktop\Research\HFSS Scripting\3dmodeler\SBend.aedt", _
    true

oDesign.Solve _
    Array("Setup1") 

Set oModule = oDesign.GetModule("Solutions")
oModule.ExportNetworkData _
        "", _
        Array("Setup1:Sweep1"), _
        7,  _
        "C:\Users\bstac\Desktop\Research\HFSS Scripting\3dmodeler\SBendData.m", _
        Array("All"),  _
        true, _
        50.00 
