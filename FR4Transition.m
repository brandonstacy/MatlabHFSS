%FR4 Transitionfunction Sdata = FR4toLiNbO3(lnbWpad, lnbLtap,Nfig); 
%Design Straight Through on FR4
close all;
clc;
clear all;

%==========================================================================
% Constants
%==========================================================================
CloseAfter = 1;
HfssSim = true;
Units = 'mm'; 
GHz = 1e9; 
Zn = 50;               % normalized impedance 
Z0 = 50;

% frequency range 
Freq0   = 40*GHz; 
Freqmin = 1*GHz;
Freqmax = 60*GHz;
nPoints = 120;

%Define the material color
GoldColor = [250 250 0];
LNBColor = [30, 70, 200];
ViaColor = [0, 0, 0];

%==========================================================================
% Design Variables
%==========================================================================
%Substrate
SubstrateMaterial = 'FR4_epoxy';
Hsub = .2;
LFR4 = 2;
Wunit = 2;

%Microstrip
MicrostripMaterial = 'copper';
HmFR4 = .017; %Half oz copper
W0FR4 = .26;
WbaseFR4 = .15; %Base Width
LbaseFR4 = .5; %Base Length
LtaperFR4 = .7; %Taper Length
GNDGapFR4 = .075;
GNDGapBaseFR4 = .075;

%Vias
RVia = .15;
ViaSpacing = .6;
EdgeSpace = .2;
Nsides = 6;
NVias = 3;

%Wirebond
WBondPad = .025;
LBondPad = .05;
HBond = .025;
LBondStraight = .01;
BondHeight = .025;
BondSpacing = .5;

%Airbox
Hairbox = 1.5;

%Waveport
Wp1 = .75;
Lp1 = .75;

Wp2 = .6;
Lp2 = .6;

%==========================================================================
% Design Shapes
%==========================================================================
%Substrate and GND
SubX = [-Wunit/2 -Wunit/2 Wunit/2 Wunit/2 -Wunit/2];
SubY  = [0 LFR4 LFR4 0 0];

%Top Ground LNB and FR4
xmg1 = WbaseFR4/2+GNDGapBaseFR4;
xmg2 = W0FR4/2 + GNDGapFR4;
ymg1 = LFR4-LbaseFR4;
ymg2 = LFR4-LbaseFR4-LtaperFR4;

MicroGNDX = [Wunit/2 Wunit/2 xmg1 xmg1 xmg2 xmg2 Wunit/2];
MicroGNDY = [0 LFR4 LFR4 ymg1 ymg2 0 0];

MicroGND1 = [MicroGNDX;MicroGNDY;zeros(length(MicroGNDX))]';
MicroGND2 = [-MicroGNDX;MicroGNDY;zeros(length(MicroGNDX))]';

%Signal LNB and FR4
MicroX = [-W0FR4/2 -W0FR4/2 -WbaseFR4/2 -WbaseFR4/2 WbaseFR4/2 WbaseFR4/2 ...
    W0FR4/2 W0FR4/2 -W0FR4/2];
MicroY = [0 ymg2 ymg1 LFR4 LFR4 ymg1 ymg2 0 0];

MicroSignal = [MicroX;MicroY;zeros(size(MicroY))]';

%Vias
for I = 1:NVias
   ViaX(I) = W0FR4/2+GNDGapFR4+EdgeSpace;
   ViaY(I) = ViaSpacing+(I-1)*ViaSpacing;
end

figure(2); 
clf;
axis equal tight;
hold on;
axis off;
% title('FR4 to LiNbO3');
% xlabel('Width (mm)');
% ylabel('Length (mm)');
patch('xdata',SubX, 'ydata',SubY,'facecolor','c'); 
patch('xdata',MicroX, 'ydata',MicroY,'facecolor','y'); 
patch('xdata',MicroGNDX, 'ydata',MicroGNDY,'facecolor','y'); 
patch('xdata',-MicroGNDX, 'ydata',MicroGNDY,'facecolor','y'); 

for I = 1:length(ViaX)
    plot(ViaX(I),ViaY(I),'-o','MarkerSize',4,'MarkerEdgeColor','black','MarkerFaceColor','white');
    plot(-ViaX(I),ViaY(I),'-o','MarkerSize',4,'MarkerEdgeColor','black','MarkerFaceColor','white');
end

% figure(2)
% clf;
% axis equal tight;
% patch('xdata',WirebondX, 'ydata',WirebondY,'facecolor','y');

if HfssSim
    %==========================================================================
    % Begin HFSS
    %==========================================================================
    % Temporary Files
    tmpPrjFile    = [pwd, '\FR4_Transition.aedt'];
    tmpDataFile   = [pwd, '\FR4_Transition_Data.m'];
    tmpScriptFile = 'FR4_Transition.vbs';
    
    %Configure Path and Delete files
    fclose('all');
    delete 'FR4_Transition.log';
    delete 'FR4_Transition.vbs';
    myCD = ['C:\Users\bstac\Desktop\Research\HFSS Scripting\'];
    addpath(genpath((myCD)));

    % HFSS Executable Path.
    hfssExePath = 'C:\"Program Files"\AnsysEM\AnsysEM19.2\Win64\ansysedt.exe';
    
    %==========================================================================
    % Create script file.
    %==========================================================================
    fid = fopen(tmpScriptFile, 'wt');
    hfssNewProject(fid);
    hfssInsertDesign(fid, 'FR4_Transition');
    
    %==========================================================================
    % Design
    %==========================================================================
    %Airbox
    hfssBox(fid, 'AirBox', [-Wunit/2,0 , 0], [Wunit,LFR4,Hairbox], Units);
    hfssAssignMaterial(fid, 'AirBox', 'air');
    hfssSetTransparency(fid, {'AirBox'}, 1);
    hfssAssignRadiation(fid, 'RadBox', 'AirBox');
     
    %Microstrip Bottom Ground
    hfssBox(fid, 'GND', [-Wunit/2,0 , 0], [Wunit,LFR4,HmFR4], Units);
    hfssAssignMaterial(fid, 'GND', MicrostripMaterial);
    hfssSetColor(fid, 'GND', GoldColor);
    hfssSetTransparency(fid, {'GND'}, 0);
    
    %Substrate
    hfssBox(fid, 'Substrate', [-Wunit/2,0 , HmFR4], [Wunit,LFR4,Hsub], Units);
    hfssAssignMaterial(fid, 'Substrate', SubstrateMaterial);
    hfssSetColor(fid, 'Substrate', LNBColor);
    hfssSetTransparency(fid, {'Substrate'}, 0.5);
    
    %Microstrip Trace
    hfssPolygon(fid,  'MicroSignal', MicroSignal, Units);
    hfssThickenSheet(fid, {'MicroSignal'}, -HmFR4, Units);
    hfssSetColor(fid, 'MicroSignal', GoldColor);
    hfssSetTransparency(fid,  {'MicroSignal'}, 0);
    hfssMove(fid,{'MicroSignal'},[0,0,Hsub+HmFR4],Units);
    hfssAssignMaterial(fid, 'MicroSignal', MicrostripMaterial);

    %Microstrip Ground
    hfssPolygon(fid,  'MicroGND1', MicroGND1, Units);
    hfssThickenSheet(fid, {'MicroGND1'}, HmFR4, Units);
    hfssSetColor(fid, 'MicroGND1', GoldColor);
    hfssSetTransparency(fid,  {'MicroGND1'}, 0);
    hfssMove(fid,{'MicroGND1'},[0,0,Hsub+HmFR4],Units);
    hfssAssignMaterial(fid, 'MicroGND1', MicrostripMaterial);
    
    hfssPolygon(fid,  'MicroGND2', MicroGND2, Units);
    hfssThickenSheet(fid, {'MicroGND2'}, -HmFR4, Units);
    hfssSetColor(fid, 'MicroGND2', GoldColor);
    hfssSetTransparency(fid,  {'MicroGND2'}, 0);
    hfssMove(fid,{'MicroGND2'},[0,0,Hsub+HmFR4],Units);
    hfssAssignMaterial(fid, 'MicroGND2', MicrostripMaterial);
    
    %Vias
    for I = 1:length(ViaX)
        hfssCylinder2(fid, "ViaSub"+num2str(I), 'Z', [ViaX(I), ViaY(I), HmFR4], RVia, Hsub, Units, Nsides);
        hfssCylinder2(fid, "ViaSub"+num2str(I+length(ViaX)), 'Z', [-ViaX(I), ViaY(I), HmFR4], RVia, Hsub+HmFR4*2, Units, Nsides);
        hfssCylinder2(fid, "Via"+num2str(I), 'Z', [ViaX(I), ViaY(I), 0], RVia, Hsub+HmFR4*2, Units, Nsides);
        hfssCylinder2(fid, "Via"+num2str(I+length(ViaX)), 'Z', [-ViaX(I), ViaY(I), 0], RVia, Hsub+HmFR4*2, Units, Nsides);
    end
    for I = 1:length(ViaX)
        hfssSubtract(fid, {'Substrate'}, {"ViaSub"+num2str(I)});
        hfssSubtract(fid, {'Substrate'}, {"ViaSub"+num2str(I+length(ViaX))});
    end
    hfssUnite(fid, {'GND', "Via"+num2str(1),"Via"+num2str(1+length(ViaX)), ...
        'MicroGND2','MicroGND1'});
    for I = 2:length(ViaX)
        hfssUnite(fid, {'GND', "Via"+num2str(I),"Via"+num2str(I+length(ViaX))});
    end
    
    %Waveports
    hfssRectangle(fid, 'Port1','Y',[-0.5*Wp1, 0, Hsub-Lp1/2], Lp1, Wp1, Units);
    hfssAssignWavePort(fid, 'WavePort1','Port1', 1, false, [0,0,0], [0,Lp1,0], Units);
    
    hfssRectangle(fid, 'Port2','Y',[-0.5*Wp2, LFR4, Hsub-Lp2/2], Lp2, Wp2, Units);
    hfssAssignWavePort(fid, 'WavePort2','Port2', 1, false, [0,0,0], [0,Lp2,0], Units);
    %==========================================================================
    % Solution
    %==========================================================================
    % Add a Solution Setup.
    hfssInsertSolution(fid, 'Setup1', Freq0/GHz);
    hfssInterpolatingSweep(fid,'Sweep1','Setup1',Freqmin/GHz,Freqmax/GHz,nPoints);
    
    % Save the project to a temporary file and solve it.
    hfssSaveProject(fid, tmpPrjFile, true);
    hfssSolveSetup(fid, 'Setup1');
    
    % Export the Network data as an m-file.
    hfssExportNetworkData(fid, tmpDataFile, 'Setup1','Sweep1','m',Zn);
    
    fclose(fid);
    %==========================================================================
    % PLOTTING
    %==========================================================================
    %Execute the Script by starting HFSS.
    disp('Solving using HFSS...');
    if (CloseAfter == 1)
        hfssExecuteScript(hfssExePath, tmpScriptFile, 0, 'true');
    else
        hfssExecuteScript(hfssExePath, tmpScriptFile, 0, 'false');
    end
    disp('Solution Completed.');
    
    % Load the data by running the exported matlab file.
    run(tmpDataFile);
    Freq = f/1e9;
    VSWR = (1 + abs(S))./(1 - abs(S)); 
    S11 = 20*log10(abs(S(:,1,1)));
    S12 = 20*log10(abs(S(:,1,2)));
    S21 = 20*log10(abs(S(:,2,1)));
    S22 = 20*log10(abs(S(:,2,2)));
    
    Sdata.Freq = Freq;
    Sdata.S11 = S11;
    Sdata.S12 = S12;
    Sdata.S21 = S21;
    Sdata.S22 = S22;
    
    %=========================================================================
    %=========================================================================
    
    % plot the results
    figure(1);
    clf;
    plot(Freq, S11, Freq, S12, Freq, S21, Freq, S22,'LineWidth',2)
    xlabel('Frequency (GHz)');
    ylabel('Return Loss (dB)');
    axis([min(Freq) max(Freq) -30 5])
    legend('S11','S12','S21','S22','Location','Best')
    box on;
    drawnow;
end