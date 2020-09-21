function hfssSetupAntennaArray(fid, NoU, NoV, Px, Py, Units, Thetas, Phis)

% Preamble.
fprintf(fid, '\n');
fprintf(fid, 'Set oModule = oDesign.GetModule("RadField") \n');
fprintf(fid, 'oModule.EditAntennaArraySetup _\n');

% Box Parameters.
fprintf(fid, 'Array("NAME:ArraySetupInfo", _\n');
fprintf(fid, '"UseOption:=", "RegularArray", _\n');
fprintf(fid, '  Array("NAME:RegularArray", _\n');

fprintf(fid, '    "NumUCells:=", "%i", "NumVCells:=", "%i", _\n', NoU, NoV);
fprintf(fid, '    "CellUDist:=", "%f%s", "CellVDist:=", "%f%s", _\n', Px,Units,Py,Units);

fprintf(fid, '    "UDirnX:=", "1", "UDirnY:=", "0", "UDirnZ:=", "0", _\n');
fprintf(fid, '    "VDirnX:=", "0", "VDirnY:=", "1", "VDirnZ:=", "0", _\n');

fprintf(fid, '    "FirstCellPosX:=", "%f%s", _\n', 0,Units);
fprintf(fid, '    "FirstCellPosY:=", "%f%s", _\n', 0,Units);
fprintf(fid, '    "FirstCellPosZ:=", "%f%s", _\n', 0,Units);
fprintf(fid, '    "UseScanAngle:=", true, _\n');

fprintf(fid, '    "ScanAnglePhi:=", "%fdeg", _\n', Phis);
fprintf(fid, '    "ScanAngleTheta:=", "%fdeg")) \n', Thetas);









