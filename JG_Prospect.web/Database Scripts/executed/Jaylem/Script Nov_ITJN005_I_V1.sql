GO
/****** Object:  StoredProcedure [dbo].[UDP_GETInstallUserDetails]    Script Date: 09-12-2016 10:38:50 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
ALTER PROCEDURE [dbo].[UDP_GETInstallUserDetails]
	@id int
As 
BEGIN

	SELECT Id,FristName,Lastname,Email,[Address],Designation,
	[Status],[Password],Phone,Picture,Attachements,zip,[state],city,
	Bussinessname,SSN,SSN1,SSN2,[Signature],DOB,Citizenship,' ',
	EIN1,EIN2,A,B,C,D,E,F,G,H,[5],[6],[7],maritalstatus,PrimeryTradeId,SecondoryTradeId,Source,Notes,StatusReason,GeneralLiability,PCLiscense,WorkerComp,HireDate,TerminitionDate,WorkersCompCode,NextReviewDate,EmpType,LastReviewDate,PayRates,ExtraEarning,ExtraEarningAmt,PayMethod,Deduction,DeductionType,AbaAccountNo,AccountNo,AccountType,PTradeOthers,
	STradeOthers,DeductionReason,InstallId,SuiteAptRoom,FullTimePosition,ContractorsBuilderOwner,MajorTools,DrugTest,ValidLicense,TruckTools,PrevApply,LicenseStatus,CrimeStatus,StartDate,SalaryReq,Avialability,ResumePath,skillassessmentstatus,assessmentPath,WarrentyPolicy,CirtificationTraining,businessYrs,underPresentComp,websiteaddress,PersonName,PersonType,CompanyPrinciple,UserType,Email2,Phone2,CompanyName,SourceUser,DateSourced,InstallerType,BusinessType,CEO,LegalOfficer,President,Owner,AllParteners,MailingAddress,Warrantyguarantee,WarrantyYrs,MinorityBussiness,WomensEnterprise,InterviewTime,CruntEmployement,CurrentEmoPlace,LeavingReason,CompLit,FELONY,shortterm,LongTerm,BestCandidate,TalentVenue,Boardsites,NonTraditional,ConSalTraning,BestTradeOne,BestTradeTwo,BestTradeThree
	,aOne,aOneTwo,bOne,cOne,aTwo,aTwoTwo,bTwo,cTwo,aThree,aThreeTwo,bThree,cThree,TC,ExtraIncomeType,RejectionDate ,UserInstallId,PositionAppliedFor
	FROM tblInstallUsers 
	WHERE ID=@id

END






