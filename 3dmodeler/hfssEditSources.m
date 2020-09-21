% ----------------------------------------------------------------------------
% function hfssSphere(fid, Name, Center, Radius, Units)
% 
% Description :
% -------------
% Creates the VB Script necessary to create a sphere in HFSS.
%
% Parameters :
% ------------
% fid     - file identifier of the HFSS script file.
% Name    - name of the sphere (appears in HFSS).
% Center  - center of the sphere (specify as [cx, cy, cz])
% Radius  - radius of the sphere (scalar)
% Units   - specify as 'in', 'm', 'meter' or anything defined in HFSS.
% 
% Note :
% ------
%
% Example :
% ---------
% fid = fopen('myantenna.vbs', 'wt');
% ... 
% hfssSphere(fid, 'Sphere1', [0,0,0], 1, 'meter');
%
% ----------------------------------------------------------------------------

% ----------------------------------------------------------------------------
% This file is part of HFSS-MATLAB-API.
%
% HFSS-MATLAB-API is free software; you can redistribute it and/or modify it 
% under the terms of the GNU General Public License as published by the Free 
% Software Foundation; either version 2 of the License, or (at your option) 
% any later version.
%
% HFSS-MATLAB-API is distributed in the hope that it will be useful, but 
% WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY 
% or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License 
% for more details.
%
% You should have received a copy of the GNU General Public License along with
% Foobar; if not, write to the Free Software Foundation, Inc., 59 Temple 
% Place, Suite 330, Boston, MA  02111-1307  USA
%
% Copyright 2004, Vijay Ramasami (rvc@ku.edu)
% ----------------------------------------------------------------------------
function hfssEditSources(fid, FieldType, SourceNames, Modes, Magnitudes, Phases, Terminated, Impedances)

% Preamble
fprintf(fid, '\n');
fprintf(fid, 'Set oModule = oDesign.GetModule("Solutions")\n');
fprintf(fid, 'oModule.EditSources ');

fprintf(fid, '"%s", _\n', FieldType);
fprintf(fid, 'Array("NAME:SourceNames", "%s"), _\n', SourceNames);
fprintf(fid, 'Array("NAME:Modes"), _\n');
fprintf(fid, 'Array("NAME:Magnitudes", %f), _\n', Magnitudes);
fprintf(fid, 'Array("NAME:Phases"), _\n');
fprintf(fid, 'Array("NAME:Terminated"), _\n');
fprintf(fid, 'Array("NAME:Impedances") _\n');


