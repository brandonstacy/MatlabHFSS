function hfssAssignPML(fid,PMLObjName,PML_Thickness,Hz,MinFreq)

RadDist = Hz/2;
UseFreq = 'true';

fprintf(fid, 'Set oModule = oDesign.GetModule("BoundarySetup")\n');
fprintf(fid, 'oModule.CreatePML _\n');
fprintf(fid, 'Array("UserDrawnGroup:=", false, _\n');
fprintf(fid, '"PMLFaces:=", Array("%d"), _\n', PMLObjName);
fprintf(fid, '"CreateJoiningObjs:=", %s, _\n', UseFreq);
fprintf(fid, '"Thickness:=", "%fmm", _\n', PML_Thickness);
fprintf(fid, '"RadDist:=", "%fmm", _\n', RadDist);
fprintf(fid, '"UseFreq:=", %s, _\n', UseFreq);
fprintf(fid, '"MinFreq:=", "%fGHz")\n', MinFreq);

