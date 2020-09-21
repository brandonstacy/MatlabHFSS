% ----------------------------------------------------------------------------
% function hfssAssignImpedance(fid, Name, Resistance, Reactance, InfGroundPlane, Objects, Faces)
% 
% Creates the VB Script necessary to assign the radiation boundary condition
% to a (closed) Object.
%
% Parameters :
% ------------
% fid     - file identifier of the HFSS script file.
% Name    - name of the impedance boundary condition (under HFSS).
% Object  - object to which the impedance boundary conditions needs to be 
%           applied.
% Resistance
% Reactance

function hfssAssignImpedance(fid, Name, Objects, Resistance, Reactance, InfGroundPlane)

fprintf(fid, '\n');

% arguments processor.
if (nargin < 5)
	error('Insufficient number of arguments!');
elseif (nargin < 6)
	InfGroundPlane = false;
end

% process UseIE
if InfGroundPlane
    InfGroundPlane = 'true';
else
    InfGroundPlane = 'false';
end

fprintf(fid, 'Set oModule = oDesign.GetModule("BoundarySetup")\n');
fprintf(fid, 'oModule.AssignImpedance ');
fprintf(fid, 'Array("NAME:%s",_\n', Name);
fprintf(fid, '"Resistance:=", "%s",_\n', Resistance);
fprintf(fid, '"Reactance:=", "%s",_\n', Reactance);

if InfGroundPlane
    fprintf(fid, '"InfGroundPlane:=", %s,_\n',InfGroundPlane);
else
    fprintf(fid, '"InfGroundPlane:=", %s_\n',InfGroundPlane);
end

fprintf(fid, '"Objects:=", Array("%s")) \n', Objects);
%fprintf(fid, '"Faces:=", "%s")\n', Faces);

%fprintf(fid, '\n');