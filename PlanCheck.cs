using System;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Media;
using System.Windows.Controls;
using System.IO;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;

// Since   :  2015/04/10  by wakita
// Updated :  2015/04/13  by wakita	
//         :  2015/04/24  by wakita
//         :  2015/05/07  by wakita
//            Change threshold value of detecting MLC Y Min/Max Edge from 0.01 mm to 0.3 mm.
//            If MLC opening is larger than 0.3 mm, that position is recognized to MLC Y Min/Max Edge.
//         :  2015/05/08  by wakita
//            Change interface from MessageBox to Window.
//         :  2015/05/22  by wakita
//            Add the function CheckRefPoint. This function checks the location of reference point
//            if the point is too close to field edge. The threshold is 5 mm.
//         :  2015/06/01  by wakita
//            Add the function to check the electron plan. SSD, field normalization method and field weight will be verified.
//         :  2015/06/03  by wakita
//            Add the function to check the HU value of Primary reference point.
//            If the HU value < -300, the warning message will pop-up.
//         :  2015/06/08  by wakita
//            Applicable to plan sum.
//         :  2015/08/03  by wakita
//            Add the function to check the CT image used in plan is the latest.
//         :  2016/02/19 by wakita
//            Change UI for checking PlanSum to be able to select Course ID.
//         :  2016/03/08 by wakita
//            Add checking if the prescribed dose and the Primary Ref Point dose are the same.
//         :  2016/06/24 by wakita
//            Add Print button to print plan report using external vbs script.
//         :  2016/06/29 by wakita
//            Add checking if the field weight is not zero.
//         :  2016/10/20 by wakita
//            Add checking if patient orientation is Head First Supine.
//         :  2017/02/13 by wakita
//            Add the function to check the ID of CT image contains the correct date in a format of "YYMMDD".

// Do not change namespace and class name
// otherwise Eclipse will not be able to run the script.
namespace VMS.TPS
{

  class Script
  {
    public Script()
    {
    }

    private struct Pair
    {
      public string Message;
      public string Type;

      public Pair(string message, string type)
      {
	Message = message;
	Type = type;
      }
    }

    private Tuple<float,float> GetLeafEdge(float x)
    {
      float x_min, x_max; 
      if(x<=10){
	x_min=-200+(x-1)*10;
	x_max=x_min+10;
      }else if(x<=50){
	x_min=-100+(x-11)*5;
	x_max=x_min+5;
      }else{
	x_min=100+(x-51)*10;
	x_max=x_min+10;
      }
      return Tuple.Create(x_min,x_max);
    }

    private int CheckRefPoint(Beam beam, VVector refpoint)
    {
      ControlPoint first = beam.ControlPoints.ElementAt(0);
      double gantry = first.GantryAngle*Math.PI/180;
      double coll = first.CollimatorAngle*Math.PI/180;
      double couch = first.PatientSupportAngle*Math.PI/180;
      VVector initial = refpoint - beam.IsocenterPosition;
      VVector converted1 = refpoint;
      VVector converted2 = refpoint;

      // Rotate the couch and Gantry
      converted1.x = (initial.x*Math.Cos(couch)-initial.z*Math.Sin(couch))*Math.Cos(gantry) + initial.y*Math.Sin(gantry);
      converted1.y = (-initial.x*Math.Cos(couch)+initial.z*Math.Sin(couch))*Math.Sin(gantry) + initial.y*Math.Cos(gantry);
      converted1.z = initial.x*Math.Sin(couch) + initial.z*Math.Cos(couch);

      // Next, rotate the collimator and set STD = 100 cm
      double scaling = 1000/(1000+converted1.y);
      converted2.x = (converted1.x*Math.Cos(coll) + converted1.z*Math.Sin(coll))*scaling;
      converted2.z = (-converted1.x*Math.Sin(coll) + converted1.z*Math.Cos(coll))*scaling;

      // Check the distance from Field Edge
      var jaws = first.JawPositions;
      var leaves = first.LeafPositions;
      if((converted2.x-jaws.X1 < 0) || (jaws.X2-converted2.x <0) || (converted2.z-jaws.Y1 < 0) || (jaws.Y2-converted2.z < 0))
	return 2;
      else if ((converted2.x-jaws.X1 < 5) || (jaws.X2-converted2.x < 5) || (converted2.z-jaws.Y1 < 5) || (jaws.Y2-converted2.z < 5))
	return 1;
      else if(beam.MLC != null)
      {
	if(converted2.z >= -195 && converted2.z <-190)
	{
	  if(converted2.x-leaves[0,0]<5 || leaves[1,0]-converted2.x<5)
	    return 1;
	  if(converted2.x-leaves[0,1]<0 || leaves[1,1]-converted2.x<0)
	    return 1;
	}else if(converted2.z>=-190 && converted2.z<-100)
	{
	  int leaf = (int)Math.Floor(converted2.z/10)+20;
	  if(converted2.x-leaves[0,leaf]<5 || leaves[1,leaf]-converted2.x<5)
	    return 1;
	  if((converted2.z>leaf*10-195) && (converted2.x-leaves[0,leaf+1]<0 || leaves[1,leaf+1]-converted2.x<0))
	    return 1;
	  if((converted2.z<leaf*10-195) && (converted2.x-leaves[0,leaf-1]<0 || leaves[1,leaf-1]-converted2.x<0))
	    return 1;
	}else if(converted2.z>=-100 && converted2.z<100)
	{
	  int leaf = (int)Math.Floor(converted2.z/5)+30;
	  if(converted2.x-leaves[0,leaf]<5 || leaves[1,leaf]-converted2.x<5)
	    return 1;
	  if(converted2.x-leaves[0,leaf+1]<0 || leaves[1,leaf+1]-converted2.x<0)
	    return 1;
	  if(converted2.x-leaves[0,leaf-1]<0 || leaves[1,leaf-1]-converted2.x<0)
	    return 1;
	}else if(converted2.z>=100 && converted2.z<190)
	{
	  int leaf = (int)Math.Floor(converted2.z/10)+40;
	  if(converted2.x-leaves[0,leaf]<5 || leaves[1,leaf]-converted2.x<5)
	    return 1;
	  if((converted2.z>(leaf-50)*10+105) && (converted2.x-leaves[0,leaf+1]<0 || leaves[1,leaf+1]-converted2.x<0))
	    return 1;
	  if((converted2.z<(leaf-50)*10+105) && (converted2.x-leaves[0,leaf-1]<0 || leaves[1,leaf-1]-converted2.x<0))
	    return 1;
	}else if(converted2.z>=190 && converted2.z<=195)
	{
	  if(converted2.x-leaves[0,59]<5 || leaves[1,59]-converted2.x<5)
	    return 1;
	  if(converted2.x-leaves[0,58]<0 || leaves[1,58]-converted2.x<0)
	    return 1;
	}
      }
      return 0;
    }

    Window planSumWindow = new Window();

    public void Execute(ScriptContext context/*, Window window*/)
    {
      Patient patient = context.Patient;
      PlanSetup planSetup = context.PlanSetup;
      PlanSum planSum = context.PlanSumsInScope.FirstOrDefault();
      if(planSetup == null && planSum == null)
      {
	throw new ApplicationException("***********ERROR***********\nPlease open the plan to be checked.\n***************************");
      }
      else if(planSetup != null)
      {
	CheckPlan(patient, planSetup, false);
      }
      else if(planSetup == null && planSum != null) 
      {
	planSumWindow.MinWidth = 200;
	planSumWindow.MinHeight = 100;
	planSumWindow.Title = "Select PlanSum";
	planSumWindow.SizeToContent = SizeToContent.WidthAndHeight;
	planSumWindow.WindowStartupLocation = WindowStartupLocation.CenterScreen;
	InitializeUI(planSumWindow, patient);
      }
    }

    private void CheckPlan(Patient patient, PlanSetup planSetup, bool IsPlanSum)
    {
      // If there's no selected plan with calculated dose throw an exception
      if (planSetup == null || planSetup.Dose == null)
      {
	if(IsPlanSum) planSumWindow.Close();
	throw new ApplicationException("*************ERROR*************\nThis plan has not been calculated!!!\n*******************************");
      }

      // Retrieve StructureSet
      StructureSet structureSet = planSetup.StructureSet;
      if (structureSet == null)
      {
	if(IsPlanSum) planSumWindow.Close();
	throw new ApplicationException("*************ERROR*************\nThe selected plan does not reference a StructureSet.\n*******************************");
      }

      // Retrieve image
      VMS.TPS.Common.Model.API.Image image = structureSet.Image;

      // Define the list of output
      List<Pair> results = new List<Pair>();

      // Flags
      bool fIsCTLatest = true;
      bool fWithCouch = false;
      bool fMinY = false;
      bool fMaxY = false;
      bool fField = true;
      bool fIsICSame = true;
      bool fHasRefPoint = false;
      bool fHasPrimaryRP = false;
      bool fHasBolus = false;
      bool fLinkBolus = true;
      bool fAllOK = true;
      bool fIsVMAT = false;
      bool fIsIMRT = false;
      bool fIsElectron = false;
      bool fIsSBRT = false;
      bool fVolumePre = false;

      // Check Patient Orientation
      PatientOrientation patientOrientation = planSetup.TreatmentOrientation;
      if((int)patientOrientation != 1)
      {
	results.Add(new Pair("***********WARNING************\nOrientation is not Head First Supine!\n*******************************", "ERROR"));
	fAllOK = false;
      }

      // Check course ID and plan ID
      string expression1 = "^Course(?<number>[0-9]{1,2}$)";
      Regex reg1 = new Regex(expression1);
      Match match1 = reg1.Match(planSetup.Course.Id);
      string number = null;
      if(!match1.Success)
	results.Add(new Pair("************ERROR*************\nCourse ID must be \"Course#\"\n*******************************", "ERROR"));
      else number = match1.Groups["number"].Value;

      string expression2 = "^(?<number>[0-9]{1,2})-[0-9]{1,2}-[0-9]{1,2}$";
      Regex reg2 = new Regex(expression2);
      Match match2 = reg2.Match(planSetup.Id);
      if(!match2.Success)
	results.Add(new Pair("************ERROR*************\nPlan ID must be \"#-#-#\"\n*******************************", "ERROR"));
      else if(number != match2.Groups["number"].Value)
	results.Add(new Pair("************ERROR*************\nFirst number of Plan ID\nmust be same as Course number\n*******************************", "ERROR"));
      
      if(IsPlanSum)
      {
	String planID = planSetup.Id;
	results.Add(new Pair(string.Format("Plan ID is...\n{0}", planID), "NORMAL"));
      }

      // Check normalization method
      string n_method = planSetup.PlanNormalizationMethod;
      if(n_method.IndexOf("Target") >= 0) fVolumePre = true;

      // Check Dose/Fraction
      Fractionation fr1 = planSetup.UniqueFractionation;
      int? NoF1 = fr1.NumberOfFractions;
      double frdose1 = fr1.PrescribedDosePerFraction.Dose;

      if (!NoF1.HasValue)
      {
	if(IsPlanSum) planSumWindow.Close();
	throw new ApplicationException("*************ERROR*************\nInput Number of Fractions!!!\n*******************************");
      }
      if (Double.IsNaN(frdose1))
      {
	if(IsPlanSum) planSumWindow.Close();
	throw new ApplicationException("*************ERROR*************\nInput Dose/Fraction!!!\n*******************************");
      }

      if (NoF1>49)
      {
	results.Add(new Pair("***********WARNING************\nNumber of fraction is too many!!\n*******************************", "ERROR"));
	fAllOK = false;
      }

      if (frdose1<100)
      {
	results.Add(new Pair("***********WARNING************\nDose per fraction is too small!!\n*******************************", "ERROR"));
	fAllOK = false;
      }

      // Check CT origin
      if(!((image.UserOrigin.x==0) && (image.UserOrigin.y==0) && (image.UserOrigin.z==0))){
	results.Add(new Pair(string.Format("*************ERROR*************\nCT origin is ({0}, {1}, {2})\n********************************", image.UserOrigin.x/10, image.UserOrigin.y/10, image.UserOrigin.z/10), "ERROR"));
	fAllOK = false;
      }

      // Check CT creation date
      var creation = (DateTime)image.Series.Study.CreationDateTime;
      var ctdate = creation.ToString("yyyyMMdd").Substring(2);
      if(image.Id.IndexOf(ctdate) < 0){
	results.Add(new Pair("***********WARNING************\nID of CT image dose not contain the\ncorrect date in a format of \"YYMMDD\".\n*******************************", "ERROR"));
	fAllOK = false;
      }else{
	results.Add(new Pair(string.Format("CT creation date is...\n{0}/{1}/{2}", creation.Year, creation.Month, creation.Day),"NORMAL"));
      }

      // Check CT image used in this plan is the latest
      var AllStructureSets = patient.StructureSets;
      DateTime datetime = DateTime.Now;
      foreach(var ss in AllStructureSets)
      {
	datetime = (DateTime)ss.Image.Series.Study.CreationDateTime;
	if(datetime > creation && ss.Image.Series.ImagingDeviceId != "" && ss.Image.Series.Study.Comment != "ARIA RadOnc Study")
	{
	  if(!(datetime.Year == creation.Year && datetime.Month == creation.Month && datetime.Day == creation.Day))
	  {
	    fIsCTLatest = false;
	    fAllOK = false;
	  }
	}
      }
      if(!fIsCTLatest)
	results.Add(new Pair("***********WARNING************\nPlanning CT image is not latest!!\n*******************************", "ERROR"));



      // Check CT to RED table
      String ct = image.Series.ImagingDeviceId;
      results.Add(new Pair(string.Format("CT to RED table is...\n{0}", ct), "NORMAL"));


      // Check plan type
      var beams = planSetup.Beams;
      int planType = 0;
      int numberOfBeams = 0;
      double weight = 0;
      bool fGetPlanType = false;
      foreach(var beam in beams){
	if(!fGetPlanType){
	  if(!beam.IsSetupField && beam.MLC != null){
	    planType = (int)beam.MLCPlanType;
	    fGetPlanType = true;
	  }else if(!beam.IsSetupField && beam.EnergyModeDisplayName.IndexOf("E") >= 0){
	    fIsElectron = true;
	    weight = beam.WeightFactor;
	  }
	}
	if(!beam.IsSetupField) numberOfBeams += 1;
      }
      if(planType == 1) fIsIMRT = true;
      else if(planType == 0 && fVolumePre) fIsSBRT = true;
      else if(planType == 3) fIsVMAT = true;

      if(fIsElectron && numberOfBeams > 1)
      {
	if(IsPlanSum) planSumWindow.Close();
	throw new ApplicationException("*************ERROR*************\nOnly 1 field is allowed in electron plan!!!\n*******************************");
      }
      if(fIsElectron && weight != 1.0)
	results.Add(new Pair("***********WARNING************\nField weight is not 1.0\nin electron plan!!!\n*******************************", "ERROR"));
      if(fIsElectron && n_method.IndexOf("No plan") < 0)
	results.Add(new Pair("***********WARNING************\nNo Plan Normalization is recommended\nin electron plan!!!\n*******************************", "ERROR"));


      // Check Machine ID
      String machine = beams.ElementAt(0).ExternalBeam.Id;
      double couchcenter = 0;
      results.Add(new Pair("Treatment room is...", "HEAD"));
      if(machine=="CL21EX")
      {
	results.Add(new Pair(string.Format("14 room"), "MACHINE"));
      }else if(machine=="Clinac iX-OPE"){
	results.Add(new Pair(string.Format("10 room"), "MACHINE"));
	// Check couch structures
	if(fIsVMAT){
	  foreach(var structure in structureSet.Structures){
	    if(structure.Id.IndexOf("CouchGrid") >= 0){
	      fWithCouch = true;
	      couchcenter = structure.CenterPoint.x/10;
	    }
	  }
	  if(!fWithCouch){
	    results.Add(new Pair(string.Format("***********WARNING************\nNo Couch Structures!!!\n*******************************"), "ERROR"));
	    fAllOK = false;
	  }else{
	    results.Add(new Pair(string.Format("x coordinate of Couch center is {0:f1} cm.", couchcenter), "NORMAL"));
	  }
	}
      }else if(machine=="CLINAC-IX"){
	results.Add(new Pair(string.Format("9 room"), "MACHINE"));
	// Check couch structures
	if(fIsVMAT){
	  foreach(var structure in structureSet.Structures){
	    if(structure.Id.IndexOf("CouchGrid") >= 0){
	      fWithCouch = true;
	      couchcenter = structure.CenterPoint.x/10;
	    }
	  }
	  if(!fWithCouch){
	    results.Add(new Pair(string.Format("***********WARNING************\nNo Couch Structures!!!\n*******************************"), "ERROR"));
	    fAllOK = false;
	  }else{
	    results.Add(new Pair(string.Format("x coordinate of Couch center is {0:f1} cm.", couchcenter), "NORMAL"));
	  }
	}
      }else if(machine=="CL-IX-13"){
	results.Add(new Pair(string.Format("13 room"), "MACHINE"));
	// Check couch structures
	if(fIsVMAT){
	  foreach(var structure in structureSet.Structures){
	    if(structure.Id.IndexOf("CouchGrid") >= 0){
	      fWithCouch = true;
	      couchcenter = structure.CenterPoint.x/10;
	    }
	  }
	  if(!fWithCouch){
	    results.Add(new Pair(string.Format("***********WARNING************\nNo Couch Structures!!!\n*******************************"), "ERROR"));
	    fAllOK = false;
	  }else{
	    results.Add(new Pair(string.Format("x coordinate of Couch center is {0:f1} cm", couchcenter), "NORMAL"));
	  }
	}
      }else if(machine=="TrueBeamSN1609"){
	results.Add(new Pair(string.Format("TrueBeam"), "MACHINE"));
	// Check couch structures
	foreach(var structure in structureSet.Structures){
	  if(structure.Id.IndexOf("CouchInterior") >= 0){
	    fWithCouch = true;
	    couchcenter = structure.CenterPoint.x/10;
	  }
	}
	if(!fWithCouch && !fIsElectron){
	  results.Add(new Pair(string.Format("***********WARNING************\nNo Couch Structures!!!\n*******************************"), "ERROR"));
	  fAllOK = false;
	}
	else if(!fIsElectron)
	{
	  results.Add(new Pair(string.Format("x coordinate of Couch center is {0:f1} cm", couchcenter), "NORMAL"));
	}
      }else{
	results.Add(new Pair(string.Format("Unknown treatment machine."), "ERROR"));
      }

      // Check Jaw/MLC position and dose rate
      float MinX, MaxX, MinY_U, MinY_L, MaxY_U, MaxY_L, MinY, MaxY;
      int i, doseRate;
      foreach(var beam in beams)
      {
	if(!beam.IsSetupField)
	{
	  results.Add(new Pair(string.Format("{0}", beam.Id), "HEAD"));
	  if(beam.MLC != null && beam.MLCPlanType == 0){
	    fField = true;
	    var jaws = beam.ControlPoints.ElementAt(0).JawPositions;
	    var leaves = beam.ControlPoints.ElementAt(0).LeafPositions;
	    MinX = 200;
	    MaxX = -200;
	    MinY = 60;
	    MaxY = 0;
	    fMinY = false;
	    fMaxY = false;
	    // Get Min X and Max X
	    for(i=0; i<60; i++){
	      if(leaves[1,i]-leaves[0,i] > 0.01){
		if(MinX > leaves[0,i]) MinX = leaves[0,i];
		if(MaxX < leaves[1,i]) MaxX = leaves[1,i];
	      }
	    }
	    // Get Min Y and Max Y
	    for(i=0; i<60; i++){
	      if((leaves[1,i]-leaves[0,i] > 0.3) && !fMinY){
		MinY = i+1;
		fMinY = true;
	      }
	      if((leaves[1,59-i]-leaves[0,59-i] > 0.3) && !fMaxY){
		MaxY = 60-i;
		fMaxY = true;
	      }
	    }
	    MinY_U = GetLeafEdge(MinY).Item2;
	    MinY_L = GetLeafEdge(MinY).Item1;
	    MaxY_U = GetLeafEdge(MaxY).Item2;
	    MaxY_L = GetLeafEdge(MaxY).Item1;

	    // Check Jaw/MLC position
	    if(Math.Abs(jaws.X1-MinX) > 2.0){
	      results.Add(new Pair(string.Format("***********WARNING************\nCheck X1 Jaw/MLC!!!\n*******************************"), "ERROR"));
	      fField = false;
	    }
	    if(Math.Abs(jaws.X2-MaxX) > 2.0){
	      results.Add(new Pair(string.Format("***********WARNING************\nCheck X2 Jaw/MLC!!!\n*******************************"), "ERROR"));
	      fField = false;
	    }
	    if((jaws.Y1-MinY_L < -2.0) || (jaws.Y1 > MinY_U)){
	      results.Add(new Pair(string.Format("***********WARNING************\nCheck Y1 Jaw/MLC!!!\n*******************************"), "ERROR"));
	      fField = false;
	    }
	    if((MaxY_U-jaws.Y2 < -2.0) || (jaws.Y2 < MaxY_L)){
	      results.Add(new Pair(string.Format("***********WARNING************\nCheck Y2 Jaw/MLC!!!\n*******************************"), "ERROR"));
	      fField = false;
	    }

	    // Check Dose Rate
	    doseRate = beam.DoseRate;
	    if(beam.EnergyModeDisplayName == "4X"){
	      if(doseRate != 250) {
		results.Add(new Pair(string.Format("***********WARNING************\nDose Rate is {0} MU/min!!\n*******************************", doseRate), "ERROR"));
		fField = false;
	      }
	    }else if(beam.EnergyModeDisplayName == "6X-FFF"){
	      if((doseRate != 1400 && !fVolumePre) || (doseRate != 1400 && fVolumePre)) {
		results.Add(new Pair(string.Format("***********WARNING************\nDose Rate is {0} MU/min!!\n*******************************", doseRate), "ERROR"));
		fField = false;
	      }
	    }else if(beam.EnergyModeDisplayName == "10X-FFF"){
	      if((doseRate != 2400 && !fVolumePre) || (doseRate != 2400 && fVolumePre)) {
		results.Add(new Pair(string.Format("***********WARNING************\nDose Rate is {0} MU/min!!\n*******************************", doseRate), "ERROR"));
		fField = false;
	      }
	    }else{
	      if((doseRate != 300 && !fVolumePre) || (doseRate != 600 && fVolumePre)) {
		results.Add(new Pair(string.Format("***********WARNING************\nDose Rate is {0} MU/min!!\n*******************************", doseRate), "ERROR"));
		fField = false;
	      }
	    }

	    // Check field weight
	    if(beam.WeightFactor == 0.0) {
	      results.Add(new Pair(string.Format("***********WARNING************\n{0} has 0 MU!!\n*******************************", beam.Id), "ERROR"));
	      fField = false;
	    }

	    if(fField) results.Add(new Pair("ok.", "HEAD"));
	    else fAllOK = false;

	  }else if(beam.MLC == null && !fIsElectron){  /* If the field has no MLC. */
	    // Check Dose Rate
	    doseRate = beam.DoseRate;
	    if(beam.EnergyModeDisplayName == "4X"){
	      if(doseRate != 250) {
		results.Add(new Pair(string.Format("***********WARNING************\nDose Rate is {0} MU/min!!\n*******************************", doseRate),"ERROR"));
		fField = false;
	      }
	    }else{
	      if(doseRate != 300) {
		results.Add(new Pair(string.Format("***********WARNING************\nDose Rate is {0} MU/min!!\n*******************************", doseRate),"ERROR"));
		fField = false;
	      }
	    }

	    // Check field weight
	    if(beam.WeightFactor == 0.0) {
	      results.Add(new Pair(string.Format("***********WARNING************\n{0} has 0 MU!!\n*******************************", beam.Id), "ERROR"));
	      fField = false;
	    }

	    if(fField) results.Add(new Pair(string.Format("ok. {0} has no MLC.", beam.Id), "HEAD"));
	    else fAllOK = false;
	  }else if(beam.MLC == null && fIsElectron){  /* If the field is electron. */
	    // Check Dose Rate
	    doseRate = beam.DoseRate;
	    if(doseRate != 300) {
	      results.Add(new Pair(string.Format("***********WARNING************\nDose Rate is {0} MU/min!!\n*******************************", doseRate),"ERROR"));
	      fField = false;
	    }

	    // Check field weight
	    if(beam.WeightFactor == 0.0) {
	      results.Add(new Pair(string.Format("***********WARNING************\n{0} has 0 MU!!\n*******************************", beam.Id), "ERROR"));
	      fField = false;
	    }

	    if(fField) results.Add(new Pair(string.Format("ok. {0} is Electron.", beam.Id), "HEAD"));
	    else fAllOK = false;
	  }else if(fIsVMAT){ // If the plan is VMAT.
	    // Check Dose Rate
	    doseRate = beam.DoseRate;
	    if(beam.EnergyModeDisplayName == "4X"){
	      if(doseRate != 250) {
		results.Add(new Pair(string.Format("***********WARNING************\nDose Rate is {0} MU/min!!\n*******************************", doseRate), "ERROR"));
		fField = false;
	      }
	    }else if(beam.EnergyModeDisplayName == "6X-FFF"){
	      if(doseRate != 1400) {
		results.Add(new Pair(string.Format("***********WARNING************\nDose Rate is {0} MU/min!!\n*******************************", doseRate), "ERROR"));
		fField = false;
	      }
	    }else if(beam.EnergyModeDisplayName == "10X-FFF"){
	      if(doseRate != 2400) {
		results.Add(new Pair(string.Format("***********WARNING************\nDose Rate is {0} MU/min!!\n*******************************", doseRate), "ERROR"));
		fField = false;
	      }
	    }else{
	      if(doseRate != 600) {
		results.Add(new Pair(string.Format("***********WARNING************\nDose Rate is {0} MU/min!!\n*******************************", doseRate), "ERROR"));
		fField = false;
	      }
	    }
	    if(fField && fIsVMAT) results.Add(new Pair(string.Format("{0} has dynamic arc MLC.", beam.Id), "HEAD"));
	    else fAllOK = false;
	  }else if(fIsIMRT){ // If the plan is IMRT.
	    // Check Dose Rate
	    doseRate = beam.DoseRate;
	    if(beam.EnergyModeDisplayName == "4X"){
	      if(doseRate != 250) {
		results.Add(new Pair(string.Format("***********WARNING************\nDose Rate is {0} MU/min!!\n*******************************", doseRate), "ERROR"));
		fField = false;
	      }
	    }else{
	      if(doseRate != 300) {
		results.Add(new Pair(string.Format("***********WARNING************\nDose Rate is {0} MU/min!!\n*******************************", doseRate), "ERROR"));
		fField = false;
	      }
	    }
	    if(fField) results.Add(new Pair(string.Format("{0} has dynamic MLC.", beam.Id), "HEAD"));
	    else fAllOK = false;
	  }
	}
      }
      results.Add(new Pair("", "HEAD"));

      // Check fraction dose and number of fractions
      results.Add(new Pair("Dose per Fraction is...", "HEAD"));
      results.Add(new Pair(string.Format("{0} cGy",frdose1), "IMPORTANT"));
      results.Add(new Pair("Number of Fraction is...", "HEAD"));
      results.Add(new Pair(string.Format("{0} fractions",NoF1), "IMPORTANT"));
      results.Add(new Pair("Total dose is...", "HEAD"));
      results.Add(new Pair(string.Format("{0} cGy",frdose1*NoF1), "IMPORTANT"));


      // Check Isocenter and ReferencePoint
      // At First, check if all isocenters are same.
      var isocenter = beams.ElementAt(0).IsocenterPosition;
      foreach(var beam in beams)
      {
	if(VVector.Distance(beam.IsocenterPosition,isocenter)!=0) fIsICSame = false;
      }
      if(!fIsICSame){
	results.Add(new Pair("***********WARNING************\nThis plan has two or more isocenters!!\n*******************************", "ERROR"));
	fAllOK = false;
      }
      // Check the CT origin is inside the BODY
      foreach(var structure in structureSet.Structures){
	if((structure.DicomType == "EXTERNAL") && !structure.IsPointInsideSegment(image.UserOrigin)){
	  results.Add(new Pair("***********WARNING************\nCT origin is outside the BODY!!\n*******************************", "ERROR"));
	  fAllOK = false;
	}
      }
      // Check the distance from IC to CT origin
      double distance = VVector.Distance(isocenter, image.UserOrigin);
      if(distance > 500){
	results.Add(new Pair("***********WARNING************\nThe distance from IC to CT origin is too far.\n*******************************", "ERROR"));
	fAllOK = false;
      }


      // Next, check if plan has PrimaryReferencePoint and Location
      bool fGetRefPoints = false;
      var refpoints = beams.ElementAt(0).FieldReferencePoints;
      if(!beams.ElementAt(0).IsSetupField) fGetRefPoints = true;
      foreach(var beam in beams){
	if(!fGetRefPoints){
	  if(!beam.IsSetupField){
	    refpoints = beam.FieldReferencePoints;
	    fGetRefPoints = true;
	  }
	}
      }
      VVector refPointLocation = isocenter;
      foreach(var refpoint in refpoints)
      {
	fHasRefPoint = true;
	if(refpoint.IsPrimaryReferencePoint){
	  fHasPrimaryRP = true;
	  refPointLocation = refpoint.RefPointLocation;
	}
      }
      if (!fHasRefPoint)
      {
	if(IsPlanSum) planSumWindow.Close();
	throw new ApplicationException("************WARNING************\nThis plan has no reference points!!\n*******************************\n\n");
      }
      if (!fHasPrimaryRP) 
      {
	if(IsPlanSum) planSumWindow.Close();
	throw new ApplicationException("************WARNING************\nThis plan has no Primary Reference Point!!\n*******************************\n\n");
      }
      if (VVector.Distance(refPointLocation, planSetup.PlanNormalizationPoint) > 0.01 && !fIsVMAT && !fIsIMRT) 
      {
	if(IsPlanSum) planSumWindow.Close();
	throw new ApplicationException("************WARNING************\nPrimary Reference Point is not\nNormalization point!!\n*******************************\n\n");
      }
      if (Math.Abs(fr1.DosePerFractionInPrimaryRefPoint.Dose-frdose1)>1e-5 && !fIsVMAT && !fIsIMRT && !fIsSBRT && !fIsElectron)
      {
	results.Add(new Pair("***********WARNING************\nCheck Plan Normalization Mode!!\n*******************************", "ERROR"));
	fAllOK = false;
      }

      // Calculate the distance of isocenter and Primary reference point
      if (Double.IsNaN(refPointLocation.x) && !fIsElectron && !fIsVMAT && !fIsIMRT && !fIsSBRT){ 
      {
	if(IsPlanSum) planSumWindow.Close();
	throw new ApplicationException("************WARNING************\nPrimary Reference Point has no location!!\n*******************************\n\n");
      }
//	fAllOK = false;
      }else if (VVector.Distance(isocenter, refPointLocation) > 0.1 && !fIsElectron && !fIsVMAT && !fIsIMRT){
	results.Add(new Pair("***********WARNING************\nIsocenter and Primary reference point\nare not the same!!\n*******************************", "ERROR"));
	fAllOK = false;
      }else if(!fIsElectron && !fIsVMAT && !fIsIMRT && !fIsSBRT){
	results.Add(new Pair("Isocenter and Primary ref point are the same.", "NORMAL"));
      }else if(!Double.IsNaN(refPointLocation.x)){
	results.Add(new Pair("***********WARNING************\nIMRT/SBRT/Electron plan does not need\nto have reference point location.\n*******************************", "ERROR"));
      }

      // Check the location of primary reference point
      if(!fIsElectron && !fIsVMAT && !fIsIMRT && !fIsSBRT)
      {
	foreach(var beam in beams)
	{
	  if(!beam.IsSetupField)
	  {
	    int result = CheckRefPoint(beam, refPointLocation);
	    if(result == 1) results.Add(new Pair(string.Format("***********WARNING************\nReference point is too close to\nfield edge of {0}!\n*******************************",beam.Id), "ERROR"));
	    if(result == 2) results.Add(new Pair(string.Format("***********WARNING************\nReference point is outside of\nfield edge of {0}!\n*******************************",beam.Id), "ERROR"));
	  }
	}
      }

      // Check the HU value of primary reference point
      if(!fIsElectron && !fIsVMAT && !fIsIMRT && !fIsSBRT)
      {
	var start = refPointLocation;
	var stop = refPointLocation;
	start.x = refPointLocation.x + 1.0;
	stop.x = refPointLocation.x - 1.0;
	double[] profile = new double[3];
	var imageprofile = image.GetImageProfile(start, stop, profile);
	if(profile[1] < -300) results.Add(new Pair("***********WARNING************\nHU value at Reference point is too low!\n*******************************", "ERROR"));
      }

      // Check Algoriths
      if(!fIsElectron) results.Add(new Pair(string.Format("Calculation Algorithm is...\n{0}", planSetup.PhotonCalculationModel), "NORMAL"));
      else results.Add(new Pair(string.Format("Calculation Algorithm is...\n{0}", planSetup.ElectronCalculationModel), "NORMAL"));
      
      // Check heterogeneity correction
      string hetero;
      bool b = planSetup.PhotonCalculationOptions.TryGetValue("HeterogeneityCorrection", out hetero);
      results.Add(new Pair(string.Format("Heterogeneity correction is {0}.", hetero), "NORMAL"));

      // Check bolus
      foreach(var structure in structureSet.Structures){
	if(structure.DicomType == "BOLUS"){
	  fHasBolus = true;
	  foreach(var beam in beams){
	    if(beam.Boluses.Count() == 0 && !beam.IsSetupField) fLinkBolus = false;
	  }
	  if (!fLinkBolus) {
	    results.Add(new Pair("***********WARNING************\nSome Fields are not linked to Bolus!!\n*******************************", "ERROR"));
	    fAllOK = false;
	  }else {
	    results.Add(new Pair("All Fields are linked to Bolus.", "NORMAL"));
	  }
	}
      }

      // Check SSD for Electron field
      if(fIsElectron)
      {
	double ssd = 0.0;
	foreach(var beam in beams)
	{
	  if(!beam.IsSetupField) ssd = beam.SSD;
	}
	if(!fHasBolus || (fHasBolus && !fLinkBolus))
	{
	  if(ssd > 1000.5 || ssd <= 999.5)
	  {
	    results.Add(new Pair("***********WARNING************\nSSD is not 100.0 cm!!\n*******************************", "ERROR"));
	    fAllOK = false;
	  }
	}else{
	  if(ssd > 1005.5 || ssd <= 1004.5)
	  {
	    results.Add(new Pair("***********WARNING************\nSSD (Bolus+) is not 100.5 cm!!\n*******************************", "ERROR"));
	    fAllOK = false;
	  }
	}
      }

      if (fAllOK) results.Add(new Pair("This plan is probably OK.", "NORMAL"));

//      MessageBox.Show(message, "Plan parameter check of " + planSetup.Id);
      if(!IsPlanSum)
	ReportResults(results, planSetup.Id);
      else
	ReportSumResults(results, planSetup.Id);
    }

    Window ResultsWindow = new Window();

    private void ReportResults(List<Pair> results, string planId)
    {
//      var window = new Window();
      var scrollView = new ScrollViewer();
      scrollView.VerticalScrollBarVisibility = ScrollBarVisibility.Auto;
      scrollView.HorizontalScrollBarVisibility = ScrollBarVisibility.Auto;
      var panel = new StackPanel();
      panel.Orientation = Orientation.Vertical;
      panel.Background = Brushes.AliceBlue;

      var header = new Label();
      header.Content = "";
      header.Margin = new Thickness(0,0,0,-15);
      panel.Children.Add(header);

      var grid = new Grid();
      grid.ColumnDefinitions.Add(new ColumnDefinition());

      // Add results
      var counter = 0;
      foreach (var item in results)
      {
	AddRow((item.Message).Replace("_","__"), item.Type, counter, grid);
	counter++;
      }

      panel.Children.Add(grid);

      var footer = new Label();
      footer.Content = "";
      panel.Children.Add(footer);

      var grid2 = new Grid();
      grid2.RowDefinitions.Add(new RowDefinition());

      var OKbutton = new Button();
      OKbutton.Content = "Cancel";
      OKbutton.FontSize = 15;
      OKbutton.IsCancel = true;
      OKbutton.IsDefault = true;
      //      OKbutton.HorizontalAlignment = HorizontalAlignment.Right;
      OKbutton.Margin = new Thickness(0,10,0,10);
      OKbutton.Width = 80;
      OKbutton.Height = 25;

      var Printbutton = new Button();
      Printbutton.Content = "Print";
      Printbutton.FontSize = 15;
      Printbutton.HorizontalAlignment = HorizontalAlignment.Right;
      Printbutton.Margin = new Thickness(0,10,0,10);
      Printbutton.Width = 80;
      Printbutton.Height = 25;
      Printbutton.Click += new RoutedEventHandler(Print_click);

      var space = new Label();
      space.Content = null;

      grid2.ColumnDefinitions.Add(new ColumnDefinition());
      grid2.Children.Add(space);
      space.SetValue(Grid.RowProperty, 0);
      space.SetValue(Grid.ColumnProperty,0);
      grid2.ColumnDefinitions.Add(new ColumnDefinition());
      grid2.Children.Add(OKbutton);
      OKbutton.SetValue(Grid.RowProperty, 0);
      OKbutton.SetValue(Grid.ColumnProperty,2);
      grid2.ColumnDefinitions.Add(new ColumnDefinition());
      grid2.Children.Add(Printbutton);
      Printbutton.SetValue(Grid.RowProperty, 0);
      Printbutton.SetValue(Grid.ColumnProperty,1);

      panel.Children.Add(grid2);

      scrollView.Content = panel;

      ResultsWindow.Content = scrollView;
      ResultsWindow.Title = "Plan Parameter Check of " + planId;
      ResultsWindow.SizeToContent = SizeToContent.WidthAndHeight;
//      window.Height = 900;
      ResultsWindow.MinWidth = 300;
      ResultsWindow.MaxHeight = 1000;
      ResultsWindow.WindowStartupLocation = WindowStartupLocation.CenterScreen;
      ResultsWindow.ShowDialog();
    }

    private void ReportSumResults(List<Pair> results, string planId)
    {
      var window = new Window();
      var scrollView = new ScrollViewer();
      scrollView.VerticalScrollBarVisibility = ScrollBarVisibility.Auto;
      scrollView.HorizontalScrollBarVisibility = ScrollBarVisibility.Auto;
      var panel = new StackPanel();
      panel.Orientation = Orientation.Vertical;
      panel.Background = Brushes.AliceBlue;

      var header = new Label();
      header.Content = "";
      header.Margin = new Thickness(0,0,0,-15);
      panel.Children.Add(header);

      var grid = new Grid();
      grid.ColumnDefinitions.Add(new ColumnDefinition());

      // Add results
      var counter = 0;
      foreach (var item in results)
      {
	AddRow((item.Message).Replace("_","__"), item.Type, counter, grid);
	counter++;
      }

      panel.Children.Add(grid);

      var footer = new Label();
      footer.Content = "";
      panel.Children.Add(footer);

      var OKbutton = new Button();
      OKbutton.Content = "OK";
      OKbutton.FontSize = 15;
      OKbutton.IsCancel = true;
      OKbutton.IsDefault = true;
      OKbutton.HorizontalAlignment = HorizontalAlignment.Right;
      OKbutton.Margin = new Thickness(0,10,30,10);
      OKbutton.Width = 80;
      OKbutton.Height = 25;

      panel.Children.Add(OKbutton);

      scrollView.Content = panel;

      window.Content = scrollView;
      window.Title = "Plan Parameter Check of " + planId;
      window.SizeToContent = SizeToContent.WidthAndHeight;
//      window.Height = 900;
      window.MinWidth = 300;
      window.MaxHeight = 1000;
      window.WindowStartupLocation = WindowStartupLocation.CenterScreen;
      window.ShowDialog();
    }

    private void AddRow(string col1, string col2, int rowIndex, Grid grid, bool isHeader = false)
    {
      grid.RowDefinitions.Add(new RowDefinition());

      var margin = col2 == "HEAD" ? new Thickness(10, 0, 0, -10) : new Thickness(10, 0, 0, 5);

      var label1 = new Label();
      label1.Content = col1;
      label1.Margin = margin;
      label1.FontSize = 13;
      if(col2 == "IMPORTANT")
      {
	label1.FontSize = 19;
	label1.FontWeight = FontWeights.Bold;
	label1.Foreground = Brushes.Blue;
      }else if(col2 == "MACHINE")
      {
	label1.FontSize = 19;
	label1.FontWeight = FontWeights.Bold;
	label1.Foreground = Brushes.Blue;
      }else if(col2 == "ERROR")
      {
	label1.FontSize = 15.5;
	label1.FontWeight = FontWeights.Bold;
	label1.Foreground = Brushes.Red;
      }
      grid.Children.Add(label1);
      label1.SetValue(Grid.RowProperty, rowIndex);
      label1.SetValue(Grid.ColumnProperty, 0);
    }

    PlanSum selectedPlanSum { get; set; }
    Patient selectedPatient { get; set; }
    Course selectedCourse { get; set; }
    ComboBox planSumCombo = new ComboBox();
    ComboBox courseCombo = new ComboBox();

    private void InitializeUI(Window window, Patient patient)
    {
      selectedPatient = patient;
      
      var panel = new StackPanel();
      panel.Orientation = Orientation.Vertical;

      var label = new Label();
      label.Content = "Select Plan Sum to be checked.";
      label.FontSize = 15;
      panel.Children.Add(label);

      var label2 = new Label();
      label2.Content = "Course ID";
      label2.FontSize = 15;
      panel.Children.Add(label2);

      courseCombo.ItemsSource = patient.Courses.Where(c=>c.PlanSums.Count()>0).OrderByDescending(c=>c.PlanSums.Count())
          .ThenByDescending(c=>c.HistoryDateTime);
      courseCombo.SelectedItem = patient.Courses.Where(c => c.PlanSums.Count() > 0).OrderBy(c=>c.HistoryDateTime).Last();
      courseCombo.MinWidth = 100;
      courseCombo.Margin = new Thickness(10, 0, 10, 0);
      courseCombo.FontSize = 15;
      courseCombo.SelectionChanged += new SelectionChangedEventHandler(Item_Changed);
      panel.Children.Add(courseCombo);

      var label3 = new Label();
      label3.Content = "Plan Sum ID";
      label3.FontSize = 15;
      panel.Children.Add(label3);

      planSumCombo.ItemsSource = patient.Courses.Where(c => c.PlanSums.Count() > 0)
          .OrderBy(c => c.HistoryDateTime).Last().PlanSums;
      planSumCombo.SelectedItem = patient.Courses.Where(c => c.PlanSums.Count() > 0)
          .OrderBy(c => c.HistoryDateTime).Last().PlanSums.Last();
      planSumCombo.MinWidth = 100;
      planSumCombo.Margin = new Thickness(10, 0, 10, 15);
      planSumCombo.FontSize = 15;
      panel.Children.Add(planSumCombo);

      var subpanel = new StackPanel();
      subpanel.Orientation = Orientation.Horizontal;

//      var nulllabel = new Label();
//      nulllabel.Content = " ";
//      nulllabel.Margin = new Thickness(0, 0, 30, 0);
//      subpanel.Children.Add(nulllabel);

      var analyzeButton = new Button();
      analyzeButton.Content = "Check";
      analyzeButton.FontSize = 15;
      analyzeButton.Margin = new Thickness(10, 0, 0, 0);
      analyzeButton.Click += new RoutedEventHandler(Button_click);
      analyzeButton.IsDefault = true;
      analyzeButton.Width = 60;
      subpanel.Children.Add(analyzeButton);

      var nulllabel2 = new Label();
      nulllabel2.Content = " ";
      subpanel.Children.Add(nulllabel2);

      var PrintButton = new Button();
      PrintButton.Content = "Print";
      PrintButton.FontSize = 15;
      PrintButton.Width = 60;
      PrintButton.Click += new RoutedEventHandler(PrintSum_click);
      subpanel.Children.Add(PrintButton);

      var nulllabel3 = new Label();
      nulllabel3.Content = " ";
      subpanel.Children.Add(nulllabel3);

      var cancelButton = new Button();
      cancelButton.Content = "Cancel";
      cancelButton.FontSize = 15;
      cancelButton.IsCancel = true;
      cancelButton.Width = 60;
      cancelButton.Margin = new Thickness(0, 0, 10, 0);
      subpanel.Children.Add(cancelButton);

      panel.Children.Add(subpanel);

      panel.Background = Brushes.Cornsilk;
      window.Content = panel;

      window.ShowDialog();
    }

    private void Item_Changed(object sender, SelectionChangedEventArgs e)
    {
        selectedCourse = courseCombo.SelectedItem as Course;
        planSumCombo.ItemsSource = selectedCourse.PlanSums;
        planSumCombo.SelectedItem = selectedCourse.PlanSums.FirstOrDefault();
    }

    private void Button_click(object sender, RoutedEventArgs e)
    {
      selectedPlanSum = planSumCombo.SelectedItem as PlanSum;
      if(selectedPlanSum == null)
      {
	MessageBox.Show("No PlanSum is selected.", "Warning");
	return;
      }
      else if(selectedPlanSum.PlanSetups.Count() < 1)
      {
	MessageBox.Show(selectedPlanSum.ToString() + " has no plans.", "Warning");
	return;
      }else
      {
	foreach(PlanSetup planSetup in selectedPlanSum.PlanSetups)
	{
	  CheckPlan(selectedPatient, planSetup, true);
	  System.Threading.Thread.Sleep(50);
	}

//	planSumWindow.Close();
      }
    }

    private void PrintSum_click(object sender, RoutedEventArgs e)
    {
      planSumWindow.Close();
      System.Diagnostics.Process.Start(@"\\172.16.212.128\Radiology-NAS\05_Software\EclipseScript\sendkey.vbs");
    }

    private void Print_click(object sender, RoutedEventArgs e)
    {
      ResultsWindow.Close();
      System.Diagnostics.Process.Start(@"\\172.16.212.128\Radiology-NAS\05_Software\EclipseScript\sendkey.vbs");
    }
  }
}
