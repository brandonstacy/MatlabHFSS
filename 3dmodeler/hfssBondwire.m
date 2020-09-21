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
function hfssBondwire(fid, Name, WireType, WireDiameter, NumSides, Pos, Dir, Distance, h, alpha, beta, WhichAxis, Units)

% Preamble
fprintf(fid, '\n');
fprintf(fid, 'oEditor.CreateBondwire _\n');
fprintf(fid, 'Array("NAME:BondwireParameters", _\n');

% Center and Radius.
fprintf(fid, '"WireType:=" , "%s", _\n' , WireType);
fprintf(fid, '"WireDiameter:=", "%f%s", _\n' , WireDiameter, Units);
fprintf(fid, '"NumSides:=", "%f", _\n' , NumSides);
fprintf(fid, '"XPadPos:=", "%f%s", _\n' , Pos(1), Units);
fprintf(fid, '"YPadPos:=", "%f%s", _\n' , Pos(2), Units);
fprintf(fid, '"ZPadPos:=", "%f%s", _\n' , Pos(3), Units);
fprintf(fid, '"XDir:=" , "%f%s", _\n', Dir(1), Units);
fprintf(fid, '"YDir:=" , "%f%s", _\n', Dir(2), Units);
fprintf(fid, '"ZDir:=" , "%f%s", _\n', Dir(3), Units);
fprintf(fid, '"Distance:=" , "%f%s", _\n', Distance, Units);
fprintf(fid, '"h1:=" , "%f%s", _\n', h(1), Units);
fprintf(fid, '"h2:=" , "%f%s", _\n', h(2), Units);
fprintf(fid, '"alpha:=" , "%fdeg", _\n', alpha);
fprintf(fid, '"beta:=" , "%fdeg", _\n', beta);
fprintf(fid, '"WhichAxis:=" , "%s"), _\n', WhichAxis);
 
% % Attributes.
fprintf(fid, 'Array("NAME:Attributes", _\n');
fprintf(fid, '"Name:=", "%s", _\n', Name);
fprintf(fid, '"Flags:=", "", _\n');
fprintf(fid, '"Color:=", "(132 132 193)", _\n');
fprintf(fid, '"Transparency:=", 0, _\n');
fprintf(fid, '"PartCoordinateSystem:=", "Global", _\n');
fprintf(fid, '"MaterialName:=", "vacuum", _\n');
fprintf(fid, '"SolveInside:=", true)\n');
fprintf(fid, '\n');
