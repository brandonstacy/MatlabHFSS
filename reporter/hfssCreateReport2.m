% ----------------------------------------------------------------------------
% function hfssCreateReport2(fid, ReportName, Type, Display, Solution,...
%                           Sweep, Context, Domain, VarObj, DataObj)
% 
% Description :
% -------------
% Creates a new report with a single trace and adds it to the Results
% branch in the project tree.
%
% Parameters :
% ------------
% fid         - file identifier of the HFSS script file.
% ReportName  - name of the report.
% Report Type - 'Modal S Parameter', 'Terminal S Parameters','Eigenmode Parameters',
%               'Fields', 'Far Fields','Near Fields', 'Emission Test'
% DisplayType - 'Rectangular Plot', 'Polar Plot', 'Radiation Pattern',
%               'Smith Chart', 'Data Table', '3D Rectangular Plot','3D Polar Plot'
% Solution    - name of the solution given to hfssInsertSolution.
% Sweep       - name of the frequency sweep given to hfssInterpolatingSweep.
% Context     - context for which the expression is being evaluated. This
%               can be an empty string if there is no context.
%               e.g. "Infinite Sphere", "Sphere", "Polyline"
% Domain      - domain for which th expression is being evaluated. This
%               can be an empty string if there is no context.
%               e.g. "Sweep" or "Time"
% VarObj      - cell array of names of the variables to be used as sweep
%               definitions for the report, including the frequency/ies.
%               The first one will be the primary sweep.
% DataObj     - TODO.
% 
%
% Example :
%--------------------------------------------------------------------------
% fid = fopen('myantenna.vbs', 'wt');
% ... create some objects here ...
% ... insert far field spheres ...
% hfssInsertFarFieldSphereSetup(fid, 'EPlaneCutSphere', [0 180 1], [90 90 1]);
% hfssInsertFarFieldSphereSetup(fid, 'HPlaneCutSphere', [90 90 1], [0 180 1]);
%
% hfssCreateReport(fid, 'E-Plane', 'Far Fields', 'Rectangular Plot', 'Setup1', ...
%                  ' ', 'EPlaceCutSphere', {'Theta','Phi','Freq'}, {'Theta','dB(DirTotal)'});
%
% hfssCreateReport(fid, 'Return Loss', 'Modal S Parameter', 'Rectangular Plot', 'Setup1',
%                 'Sweep1', [], {'Freq'}, {'Freq','db(S(LumpedPort,LumpedPort))'});
%
%
%--------------------------------------------------------------------------
%--------------------------------------------------------------------------
function hfssCreateReport2(fid, ReportName, ReportType, DisplayType, Solution,...
                          Sweep, Context, VarObj, DataObj, Freq)
                      
% Check Sweep name, if empty it will be LastAdaptive
if isempty(Sweep)
    Sweep = 'LastAdaptive';
end

% Preamble.
fprintf(fid, '\n');
fprintf(fid, 'Set oModule = oDesign.GetModule("ReportSetup")\n');
fprintf(fid, 'oModule.CreateReport "%s", _\n', ReportName);
fprintf(fid, '"%s", ', ReportType);
fprintf(fid, '"%s", ', DisplayType);
fprintf(fid, '"%s : %s", _\n', Solution, Sweep);

Domain = 'Sweep';
% Context parameters
fprintf(fid, 'Array(');
fprintf(fid, '"Context:=", "%s"', Context);
fprintf(fid, ', ');
fprintf(fid, '"Domain:=", "%s"', Domain);
fprintf(fid, '), _\n');



% Families array parameters
fprintf(fid, 'Array('); 
if numel(VarObj) > 1
    for i = 1:numel(VarObj)-1
        fprintf(fid, '"%s:=", Array("All"), ', VarObj{i});
    end
    fprintf(fid, '"%s:=", Array("%fGHz")), _\n', VarObj{i+1}, Freq);
else
    fprintf(fid, '"%s:=", Array("All")), _\n', VarObj{1});
end


fprintf(fid, 'Array(');
switch DisplayType
    case 'Rectangular Plot'
        fprintf(fid, '"X Component:=", "%s", ', DataObj{1});
        fprintf(fid, '"Y Component:=", Array("%s")), _\n', DataObj{2});
    case 'Polar Plot'
        error('Error in hfssCreateReport: display not supported');
    case 'Radiation Pattern'
        fprintf(fid, '"Ang Component:=", "%s", ', DataObj{1});
        fprintf(fid, '"Mag Component:=", Array("%s")), _\n', DataObj{2});
    case 'Smith Chart'
        error('Error in hfssCreateReport: display not supported');
    case 'Data Table'
        error('Error in hfssCreateReport: display not supported');
    case '3D Rectangular Plot'
        fprintf(fid, '"X Component:=", "%s", ', DataObj{1});
        fprintf(fid, '"Y Component:=", "%s", ', DataObj{2});
        fprintf(fid, '"Z Component:=", Array("%s")), _\n', DataObj{3});
    case '3D Polar Plot'
        fprintf(fid, '"Phi Component:=", "%s", ', DataObj{1});
        fprintf(fid, '"Theta Component:=", "%s", ', DataObj{2});
        fprintf(fid, '"Mag Component:=", Array("%s")), _\n', DataObj{3});
    otherwise
        error('Error in hfssCreateReport: DisplayType = wrong value');
end
fprintf(fid, 'Array()\n');



% Set oModule = oDesign.GetModule("ReportSetup")
% oModule.CreateReport "Rad3D", _
% "Far Fields", "3D Polar Plot", "Setup1 : LastAdaptive", _
% Array("Context:=", "EPlaneCutSphere", "Domain:=", "Sweep"), _
% Array("theta:=", Array("All"), "phi:=", Array("All"), "Freq:=", Array("12GHz")), _
% Array("Phi Component:=", "phi", "Theta Component:=", "theta", "Mag Component:=", Array("dB(GainTotal)")), _
% Array()
%
% Set oModule = oDesign.GetModule("ReportSetup")
% oModule.CreateReport "Rad3D", _
% "Far Fields", "3D Polar Plot", "Setup1 : LastAdaptive", _
% Array("Context:=", "EPlaneCutSphere", "Domain:=", "Sweep"), _
% Array("theta:=", Array("All"), "phi:=", Array("All"), "Freq:=", Array("12GHz")), _
% Array("Phi Component:=", "phi", "Theta Component:=", "theta", "Mag Component:=", Array("GainTotal")), _
% Array()
%
% Set oModule = oDesign.GetModule("ReportSetup")
% oModule.CreateReport "Rept2DRectFreq",_
% "Modal Solution Data", "Rectangular Plot", _
% "Setup1 : Sweep1",_
% Array("Domain:=", "Sweep"), _
% Array("Freq:=", Array("All")), _
% Array("X Component:=", "Freq", "Y Component:=", Array("dB(S(LumpedPort,LumpedPort))")), _
% Array()

% Set oModule = oDesign.GetModule("ReportSetup")
% oModule.CreateReport "Rad_field", "Far Fields",_
% "3D Cartesian Plot", "Setup1 : LastAdaptive", _
% Array("Context:=", "PlaneCutSphere", "Domain:=", "Sweep"), _
% Array("Theta:=", Array("All"), "Phi:=", Array("All"), _
% "Freq:=", Array("All")), _
% Array("X Component:=", "Theta", _
% "Y Component:=", "Phi", _
% "Z Component:=", Array("GainTotal")), _
% Array()
%
% Set oModule = oDesign.GetModule("ReportSetup")
% oModule.CreateReport "ReptSmithFreq",_
% "Modal Solution Data", "Smith Plot", "Setup1 : Sweep1", _
% Array("Domain:=", "Sweep"), _
% Array("Freq:=", Array("All")),_
% Array("Polar Component:=", Array("ln(Y(LumpedPort,LumpedPort))")), _
% Array()



