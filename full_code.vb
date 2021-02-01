Public Sub daoCreateTables()
 'create some tables
 Dim db As DAO.Database
 Dim tdf As DAO.TableDef
 Dim prp As DAO.Property
 
 Set db = CurrentDb
 On Error Resume Next
 
 'create the table definition
 Set tdf = db.CreateTableDef("SATU_db_1")
 
 'create the field definitions
 Dim fldID As DAO.Field
 Set fldID = tdf.CreateField("autoID", dbLong)
 fldID.Attributes = dbAutoIncrField
 fldID.Required = True
 tdf.Fields.Append fldID
 
 'add the table to the database
 db.TableDefs.Append tdf
 Set tdf = db.TableDefs("SATU_db_1")
 
 'add other fields
 Dim fldB01 As DAO.Field
 Set fldB01 = tdf.CreateField("ResearchID_1", dbText)
 tdf.Fields.Append fldB01
 
 'dem fields
 Dim fldB02 As DAO.Field
 Set fldB02 = tdf.CreateField("Gender", dbText)
 tdf.Fields.Append fldB02
 Call setComboProperties(fldB02, "Female;Male;Transgender/Transexual;Unknown")

 Dim fldB03 As DAO.Field
 Set fldB03 = tdf.CreateField("Citizen", dbText)
 tdf.Fields.Append fldB03
 Call setComboProperties(fldB03, "Yes;No;Unknown")

 Dim fldB04 As DAO.Field
 Set fldB04 = tdf.CreateField("Ethnicity", dbText)
 tdf.Fields.Append fldB04
 Call setComboProperties(fldB04, "Hispanic other;Hispanic Puerto Rican;Hispanic-Cuban;Non-Hispanic;Hispanic Mexican;Unknown")

 Dim fldB05 As DAO.Field
 Set fldB05 = tdf.CreateField("Race", dbText)
 tdf.Fields.Append fldB05
 Call setComboMultiProperties(fldB05, "American Indian/Nativan Alaskan;Asian;Black/African American;Native Hawaiian/other Pacific Islander;White/Caucasian;Other;Unknown")

 Dim fldB06 As DAO.Field
 Set fldB06 = tdf.CreateField("Religion", dbText)
 tdf.Fields.Append fldB06
 Call setComboProperties(fldB06, "Christian (non-Catholic, non-specific);Catholic;Protestant;Protestant, other;Pentecostal;Baptist;Jehovah's witness;Judaism;Mormon/Latter Day Saints;Islam/Muslim;Buddhism;Hinduism;Native American;None;Other;Unknown")

 Dim fldB07 As DAO.Field
 Set fldB07 = tdf.CreateField("English", dbText)
 tdf.Fields.Append fldB07
 Call setComboProperties(fldB07, "Excellent;Good;Moderate;Poor;Not at all")

 Dim fldB08 As DAO.Field
 Set fldB08 = tdf.CreateField("Interpreter_Needed", dbInteger)
 tdf.Fields.Append fldB08
 Call yesNoNoneQuestions(fldB08)
 '  Call SetPropertyDAO(fldB08, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldB09 As DAO.Field
 Set fldB09 = tdf.CreateField("Language", dbText)
 tdf.Fields.Append fldB09
 Call setComboProperties(fldB09, "English;Arabic;Cambodian;Cantonese;Czeck;French;German;Greek;Haitian Creole;Hindi;Hmong;Indian (general);Italian;Japanese;Korean;Laotian;Latvian;Mandarin;Polish;Portuguese;Russian;Spanish;Thai;Vietnamese;American Sign Language;Braille;Other;Unknown")

 Dim fldB10 As DAO.Field
 Set fldB10 = tdf.CreateField("Marital", dbText)
 tdf.Fields.Append fldB10
 Call setComboProperties(fldB10, "Never married;Married;Separated;Divorce/annulled;Widowed;Civil union;Other;Unknown")

 Dim fldB11 As DAO.Field
 Set fldB11 = tdf.CreateField("Veteran", dbText)
 tdf.Fields.Append fldB11
 Call setComboProperties(fldB11, "WWII Era;Korean Hostilities;Vietnam Era;Global War on Terror;Veteran-other dates;Veteran-dates undetermined;Not a veteran;Veteran status unknown")

 Dim fldB12 As DAO.Field
 Set fldB12 = tdf.CreateField("VA_connected", dbInteger)
 tdf.Fields.Append fldB12
 Call yesNoNoneQuestions(fldB12)
 '  Call SetPropertyDAO(fldB12, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldB13 As DAO.Field
 Set fldB13 = tdf.CreateField("VA_referral", dbInteger)
 tdf.Fields.Append fldB13
 Call yesNoNoneQuestions(fldB13)
 '  Call SetPropertyDAO(fldB13, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldB14 As DAO.Field
 Set fldB14 = tdf.CreateField("referral", dbText)
 tdf.Fields.Append fldB14
 Call setComboProperties(fldB14, "Self;Family/friend;Community mental health/LMHA;Other mental health provider;substance abuse provider;Mental health care clinic;Medical health practitioner;inpatient transfer;DMHAS IP;non-DMHAS IP;school;Employer/supervisor;Employee assistance program;Clergy/church/synagogue;Home for the aged;Nursing facility;Dept of Children and families;Dept of Social services;Dept of developmental services;Other community referral;court order;Probation/parole;Prison;Police;Police community relations officer;shelter;Emergency Department;Other;Unknown")

 Dim fldB15 As DAO.Field
 Set fldB15 = tdf.CreateField("Living_Situatio", dbText)
 tdf.Fields.Append fldB15
 Call setComboProperties(fldB15, "Private residence-client owns/rents;Private residence-friend, family owns/rents;Single room occupancy;Private residence-community agency;Residential care home;Congregate residential care;Crisis/respite bed;Skilled nursing facility;Psychiatric/SA/medical inpatient;Correctional Facility;Homeless shelter;Homeless (including on street);Domestic violence shelter;Other;Unknown")

 Dim fldB16 As DAO.Field
 Set fldB16 = tdf.CreateField("Education", dbText)
 tdf.Fields.Append fldB16
 Call setComboProperties(fldB16, "No schooling;Highest grade completed;High school/GED;College;Graduate School;Voc/Tech/Business School;Other;Unknown")

 Dim fldB17 As DAO.Field
 Set fldB17 = tdf.CreateField("Employment", dbText)
 tdf.Fields.Append fldB17
 Call setComboProperties(fldB17, "Employed full-time;Employed part-time;Unemployed;Paid transitional;Paid non-comp;NILF homemaker;NILF incarcerated;NILF retired;NILF SSI;NILF student;NILF other;Volunteer;Unknown")

 Dim fldB18 As DAO.Field
 Set fldB18 = tdf.CreateField("Income", dbText)
 tdf.Fields.Append fldB18
 Call setComboProperties(fldB18, "None;Public assistance;Retirement;Salary;Disability;Other;Unknown")

 Dim fldB19 As DAO.Field
 Set fldB19 = tdf.CreateField("Pregnant", dbText)
 tdf.Fields.Append fldB19
 Call setComboProperties(fldB19, "Yes;No;Unknown;None")

 Dim fldB20 As DAO.Field
 Set fldB20 = tdf.CreateField("Homeless_not_in_shelter", dbText)
 tdf.Fields.Append fldB20
 Call setComboProperties(fldB20, "Yes;No;Unknown;None")

 Dim fldB21 As DAO.Field
 Set fldB21 = tdf.CreateField("Supported_by_family", dbText)
 tdf.Fields.Append fldB21
 Call setComboProperties(fldB21, "Yes;No;Unknown;None")

 Dim fldB22 As DAO.Field
 Set fldB22 = tdf.CreateField("Jail_Hospital", dbText)
 tdf.Fields.Append fldB22
 Call setComboProperties(fldB22, "Yes;No;Unknown;None")

 Dim fldB23 As DAO.Field
 Set fldB23 = tdf.CreateField("Arrested", dbText)
 tdf.Fields.Append fldB23
 Call setComboProperties(fldB23, "Yes;No;Unknown;None")

 Dim fldB24 As DAO.Field
 Set fldB24 = tdf.CreateField("Self-help", dbText)
 tdf.Fields.Append fldB24
 Call setComboProperties(fldB24, "Yes;No;Unknown;None")

 Dim fldB25 As DAO.Field
 Set fldB25 = tdf.CreateField("self-help-num", dbInteger)
 tdf.Fields.Append fldB25

 Dim fldB26 As DAO.Field
 Set fldB26 = tdf.CreateField("Tobacco", dbText)
 tdf.Fields.Append fldB26
 Call setComboProperties(fldB26, "1-3 times in 30 days;Once a week;3-6 times a week;Daily;3-6 times a days;More than 6 times a day;NA;Unknown")

 Dim fldB27 As DAO.Field
 Set fldB27 = tdf.CreateField("Registered_voter", dbText)
 tdf.Fields.Append fldB27
 Call setComboProperties(fldB27, "Yes;No;N/A")
 
 'phq fields
 Dim fldC01 As DAO.Field
 Set fldC01 = tdf.CreateField("PHQ_1", dbBoolean)
 tdf.Fields.Append fldC01
 Call SetPropertyDAO(fldC01, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldC02 As DAO.Field
 Set fldC02 = tdf.CreateField("PHQ_2", dbBoolean)
 tdf.Fields.Append fldC02
 Call SetPropertyDAO(fldC02, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldC03 As DAO.Field
 Set fldC03 = tdf.CreateField("PHQ_3", dbBoolean)
 tdf.Fields.Append fldC03
 Call SetPropertyDAO(fldC03, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldC04 As DAO.Field
 Set fldC04 = tdf.CreateField("PHQ_4", dbBoolean)
 tdf.Fields.Append fldC04
 Call SetPropertyDAO(fldC04, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldC05 As DAO.Field
 Set fldC05 = tdf.CreateField("PHQ_5", dbBoolean)
 tdf.Fields.Append fldC05
 Call SetPropertyDAO(fldC05, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldC06 As DAO.Field
 Set fldC06 = tdf.CreateField("PHQ_6", dbBoolean)
 tdf.Fields.Append fldC06
 Call SetPropertyDAO(fldC06, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldC07 As DAO.Field
 Set fldC07 = tdf.CreateField("PHQ_7", dbBoolean)
 tdf.Fields.Append fldC07
 Call SetPropertyDAO(fldC07, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldC08 As DAO.Field
 Set fldC08 = tdf.CreateField("PHQ_8", dbBoolean)
 tdf.Fields.Append fldC08
 Call SetPropertyDAO(fldC08, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldC09 As DAO.Field
 Set fldC09 = tdf.CreateField("PHQ_9", dbBoolean)
 tdf.Fields.Append fldC09
 Call SetPropertyDAO(fldC09, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldC10 As DAO.Field
 Set fldC10 = tdf.CreateField("PHQ_imp", dbText)
 tdf.Fields.Append fldC10
 Call setComboProperties(fldC10, "Not difficult at all;Somewhat difficult;Very difficult;Extremely difficult")

 'food fields
 Dim fldD01 As DAO.Field
 Set fldD01 = tdf.CreateField("Food_Sec1", dbText)
 tdf.Fields.Append fldD01
 Call setComboProperties(fldD01, "Often true;Sometimes true;Never true;Don't know, or refused")

 Dim fldD02 As DAO.Field
 Set fldD02 = tdf.CreateField("Food_Sec2", dbText)
 tdf.Fields.Append fldD02
 Call setComboProperties(fldD02, "Often true;Sometimes true;Never true;Don't know, or refused")

 'RAD fields
 Dim fldE01 As DAO.Field
 Set fldE01 = tdf.CreateField("A_1_1", dbInteger)
 tdf.Fields.Append fldE01
 Call intOneToSevenRange(fldE01)

 Dim fldE02 As DAO.Field
 Set fldE02 = tdf.CreateField("A_2_2", dbInteger)
 tdf.Fields.Append fldE02
 Call intOneToSevenRange(fldE02)

 Dim fldE03 As DAO.Field
 Set fldE03 = tdf.CreateField("A_3_3", dbInteger)
 tdf.Fields.Append fldE03
 Call intOneToSevenRange(fldE03)

 Dim fldE04 As DAO.Field
 Set fldE04 = tdf.CreateField("A_4_4", dbInteger)
 tdf.Fields.Append fldE04
 Call intOneToSevenRange(fldE04)

 Dim fldE05 As DAO.Field
 Set fldE05 = tdf.CreateField("A_5_5", dbInteger)
 tdf.Fields.Append fldE05
 Call intOneToSevenRange(fldE05)

 Dim fldE06 As DAO.Field
 Set fldE06 = tdf.CreateField("A_imp", dbText)
 tdf.Fields.Append fldE06
 Call setComboProperties(fldE06, "Not difficult at all;Somewhat difficult;Very difficult;Extremely difficult")

 Dim fldE07 As DAO.Field
 Set fldE07 = tdf.CreateField("D_1_6", dbInteger)
 tdf.Fields.Append fldE07
 Call intOneToSevenRange(fldE07)

 Dim fldE08 As DAO.Field
 Set fldE08 = tdf.CreateField("D_2_7", dbInteger)
 tdf.Fields.Append fldE08
 Call intOneToSevenRange(fldE08)

 Dim fldE09 As DAO.Field
 Set fldE09 = tdf.CreateField("D_3_8", dbInteger)
 tdf.Fields.Append fldE09
 Call intOneToSevenRange(fldE09)

 Dim fldE10 As DAO.Field
 Set fldE10 = tdf.CreateField("D_4_9", dbInteger)
 tdf.Fields.Append fldE10
 Call intOneToSevenRange(fldE10)

 Dim fldE11 As DAO.Field
 Set fldE11 = tdf.CreateField("D_5_10", dbInteger)
 tdf.Fields.Append fldE11
 Call intOneToSevenRange(fldE11)

 Dim fldE12 As DAO.Field
 Set fldE12 = tdf.CreateField("D_imp", dbText)
 tdf.Fields.Append fldE12
 Call setComboProperties(fldE12, "Not difficult at all;Somewhat difficult;Very difficult;Extremely difficult")

 Dim fldE13 As DAO.Field
 Set fldE13 = tdf.CreateField("S_1_11", dbInteger)
 tdf.Fields.Append fldE13
 Call intOneToSevenRange(fldE13)

 Dim fldE14 As DAO.Field
 Set fldE14 = tdf.CreateField("S_2_12", dbInteger)
 tdf.Fields.Append fldE14
 Call intOneToSevenRange(fldE14)

 Dim fldE15 As DAO.Field
 Set fldE15 = tdf.CreateField("S_3_13", dbInteger)
 tdf.Fields.Append fldE15
 Call intOneToSevenRange(fldE15)

 Dim fldE16 As DAO.Field
 Set fldE16 = tdf.CreateField("S_4_14", dbInteger)
 tdf.Fields.Append fldE16
 Call intOneToSevenRange(fldE16)

 Dim fldE17 As DAO.Field
 Set fldE17 = tdf.CreateField("S_5_15", dbInteger)
 tdf.Fields.Append fldE17
 Call intOneToSevenRange(fldE17)

 Dim fldE18 As DAO.Field
 Set fldE18 = tdf.CreateField("S_imp", dbText)
 tdf.Fields.Append fldE18
 Call setComboProperties(fldE18, "Not difficult at all;Somewhat difficult;Very difficult;Extremely difficult")

 Dim fldE19 As DAO.Field
 Set fldE19 = tdf.CreateField("G_1_16", dbInteger)
 tdf.Fields.Append fldE19
 Call intOneToSevenRange(fldE19)

 Dim fldE20 As DAO.Field
 Set fldE20 = tdf.CreateField("G_2_17", dbInteger)
 tdf.Fields.Append fldE20
 Call intOneToSevenRange(fldE20)

 Dim fldE21 As DAO.Field
 Set fldE21 = tdf.CreateField("G_3_18", dbInteger)
 tdf.Fields.Append fldE21
 Call intOneToSevenRange(fldE21)

 Dim fldE22 As DAO.Field
 Set fldE22 = tdf.CreateField("G_4_19", dbInteger)
 tdf.Fields.Append fldE22
 Call intOneToSevenRange(fldE22)

 Dim fldE23 As DAO.Field
 Set fldE23 = tdf.CreateField("G_5_20", dbInteger)
 tdf.Fields.Append fldE23
 Call intOneToSevenRange(fldE23)

 Dim fldE24 As DAO.Field
 Set fldE24 = tdf.CreateField("G_imp", dbText)
 tdf.Fields.Append fldE24
 Call setComboProperties(fldE24, "Not difficult at all;Somewhat difficult;Very difficult;Extremely difficult")

 Dim fldE25 As DAO.Field
 Set fldE25 = tdf.CreateField("B_1_21", dbInteger)
 tdf.Fields.Append fldE25
 Call intOneToSevenRange(fldE25)

 Dim fldE26 As DAO.Field
 Set fldE26 = tdf.CreateField("B_2_22", dbInteger)
 tdf.Fields.Append fldE26
 Call intOneToSevenRange(fldE26)

 Dim fldE27 As DAO.Field
 Set fldE27 = tdf.CreateField("B_3_23", dbInteger)
 tdf.Fields.Append fldE27
 Call intOneToSevenRange(fldE27)

 Dim fldE28 As DAO.Field
 Set fldE28 = tdf.CreateField("B_4_24", dbInteger)
 tdf.Fields.Append fldE28
 Call intOneToSevenRange(fldE28)

 Dim fldE29 As DAO.Field
 Set fldE29 = tdf.CreateField("B_5_25", dbInteger)
 tdf.Fields.Append fldE29
 Call intOneToSevenRange(fldE29)

 Dim fldE30 As DAO.Field
 Set fldE30 = tdf.CreateField("B_imp", dbText)
 tdf.Fields.Append fldE30
 Call setComboProperties(fldE30, "Not difficult at all;Somewhat difficult;Very difficult;Extremely difficult")

 Dim fldE31 As DAO.Field
 Set fldE31 = tdf.CreateField("P_1_26", dbInteger)
 tdf.Fields.Append fldE31
 Call intOneToSevenRange(fldE31)

 Dim fldE32 As DAO.Field
 Set fldE32 = tdf.CreateField("P_2_27", dbInteger)
 tdf.Fields.Append fldE32
 Call intOneToSevenRange(fldE32)

 Dim fldE33 As DAO.Field
 Set fldE33 = tdf.CreateField("P_3_28", dbInteger)
 tdf.Fields.Append fldE33
 Call intOneToSevenRange(fldE33)

 Dim fldE34 As DAO.Field
 Set fldE34 = tdf.CreateField("P_4_29", dbInteger)
 tdf.Fields.Append fldE34
 Call intOneToSevenRange(fldE34)

 Dim fldE35 As DAO.Field
 Set fldE35 = tdf.CreateField("P_5_30", dbInteger)
 tdf.Fields.Append fldE35
 Call intOneToSevenRange(fldE35)

 Dim fldE36 As DAO.Field
 Set fldE36 = tdf.CreateField("P_imp", dbText)
 tdf.Fields.Append fldE36
 Call setComboProperties(fldE36, "Not difficult at all;Somewhat difficult;Very difficult;Extremely difficult")

 Dim fldE37 As DAO.Field
 Set fldE37 = tdf.CreateField("V_1_31", dbInteger)
 tdf.Fields.Append fldE37
 Call intOneToSevenRange(fldE37)

 Dim fldE38 As DAO.Field
 Set fldE38 = tdf.CreateField("V_2_32", dbInteger)
 tdf.Fields.Append fldE38
 Call intOneToSevenRange(fldE38)

 Dim fldE39 As DAO.Field
 Set fldE39 = tdf.CreateField("V_3_33", dbInteger)
 tdf.Fields.Append fldE39
 Call intOneToSevenRange(fldE39)

 Dim fldE40 As DAO.Field
 Set fldE40 = tdf.CreateField("V_4_34", dbInteger)
 tdf.Fields.Append fldE40
 Call intOneToSevenRange(fldE40)

 Dim fldE41 As DAO.Field
 Set fldE41 = tdf.CreateField("V_5_35", dbInteger)
 tdf.Fields.Append fldE41
 Call intOneToSevenRange(fldE41)

 Dim fldE42 As DAO.Field
 Set fldE42 = tdf.CreateField("V_imp", dbText)
 tdf.Fields.Append fldE42
 Call setComboProperties(fldE42, "Not difficult at all;Somewhat difficult;Very difficult;Extremely difficult")

 'vitals fields
 Dim fldF01 As DAO.Field
 Set fldF01 = tdf.CreateField("Vitals_Date", dbDate)
 tdf.Fields.Append fldF01

 Dim fldF02 As DAO.Field
 Set fldF02 = tdf.CreateField("BP_sys", dbInteger)
 tdf.Fields.Append fldF02

 Dim fldF03 As DAO.Field
 Set fldF03 = tdf.CreateField("BP_dia", dbInteger)
 tdf.Fields.Append fldF03

 Dim fldF04 As DAO.Field
 Set fldF04 = tdf.CreateField("Heart_Rate", dbInteger)
 tdf.Fields.Append fldF04

 Dim fldF05 As DAO.Field
 Set fldF05 = tdf.CreateField("vit_height-ft", dbInteger)
 tdf.Fields.Append fldF05

 Dim fldF06 As DAO.Field
 Set fldF06 = tdf.CreateField("vit_height-in", dbInteger)
 tdf.Fields.Append fldF06

 Dim fldF07 As DAO.Field
 Set fldF07 = tdf.CreateField("vit_weight", dbDouble)
 tdf.Fields.Append fldF07

 Dim fldF08 As DAO.Field
 Set fldF08 = tdf.CreateField("BMI", dbInteger)
 tdf.Fields.Append fldF08

 Dim fldF09 As DAO.Field
 Set fldF09 = tdf.CreateField("Current_Smoker", dbBoolean)
 tdf.Fields.Append fldF09
 Call SetPropertyDAO(fldF09, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldF10 As DAO.Field
 Set fldF10 = tdf.CreateField("IF_Smoker_offered_help", dbBoolean)
 tdf.Fields.Append fldF10
 Call SetPropertyDAO(fldF10, "DisplayControl", dbInteger, CInt(acCheckBox))

 'mental health screen fields
 Dim fldG01 As DAO.Field
 Set fldG01 = tdf.CreateField("MH_1", dbInteger)
 tdf.Fields.Append fldG01
 '  Call SetPropertyDAO(fldG01, "DisplayControl", dbInteger, CInt(acCheckBox))
 Call yesNoNoneQuestions(fldG01)

 Dim fldG02 As DAO.Field
 Set fldG02 = tdf.CreateField("MH_2", dbInteger)
 tdf.Fields.Append fldG02
 '  Call SetPropertyDAO(fldG02, "DisplayControl", dbInteger, CInt(acCheckBox))
 Call yesNoNoneQuestions(fldG02)

 Dim fldG03 As DAO.Field
 Set fldG03 = tdf.CreateField("MH_3", dbInteger)
 tdf.Fields.Append fldG03
 '  Call SetPropertyDAO(fldG03, "DisplayControl", dbInteger, CInt(acCheckBox))
 Call yesNoNoneQuestions(fldG03)

 Dim fldG04 As DAO.Field
 Set fldG04 = tdf.CreateField("MH_4", dbInteger)
 tdf.Fields.Append fldG04
 '  Call SetPropertyDAO(fldG04, "DisplayControl", dbInteger, CInt(acCheckBox))
 Call yesNoNoneQuestions(fldG04)

 Dim fldG05 As DAO.Field
 Set fldG05 = tdf.CreateField("MH_5", dbInteger)
 tdf.Fields.Append fldG05
 '  Call SetPropertyDAO(fldG05, "DisplayControl", dbInteger, CInt(acCheckBox))
 Call yesNoNoneQuestions(fldG05)

 Dim fldG06 As DAO.Field
 Set fldG06 = tdf.CreateField("MH_6a", dbInteger)
 tdf.Fields.Append fldG06
 '  Call SetPropertyDAO(fldG06, "DisplayControl", dbInteger, CInt(acCheckBox))
 Call yesNoNoneQuestions(fldG06)

 Dim fldG07 As DAO.Field
 Set fldG07 = tdf.CreateField("MH_6b", dbInteger)
 tdf.Fields.Append fldG07
 '  Call SetPropertyDAO(fldG07, "DisplayControl", dbInteger, CInt(acCheckBox))
 Call yesNoNoneQuestions(fldG07)

 Dim fldG08 As DAO.Field
 Set fldG08 = tdf.CreateField("MH_7", dbInteger)
 tdf.Fields.Append fldG08
 '  Call SetPropertyDAO(fldG08, "DisplayControl", dbInteger, CInt(acCheckBox))
 Call yesNoNoneQuestions(fldG08)

 Dim fldG09 As DAO.Field
 Set fldG09 = tdf.CreateField("MH_8", dbInteger)
 tdf.Fields.Append fldG09
 '  Call SetPropertyDAO(fldG09, "DisplayControl", dbInteger, CInt(acCheckBox))
 Call yesNoNoneQuestions(fldG09)

 Dim fldG10 As DAO.Field
 Set fldG10 = tdf.CreateField("MH_9", dbInteger)
 tdf.Fields.Append fldG10
 '  Call SetPropertyDAO(fldG10, "DisplayControl", dbInteger, CInt(acCheckBox))
 Call yesNoNoneQuestions(fldG10)

 Dim fldG11 As DAO.Field
 Set fldG11 = tdf.CreateField("MH_10", dbInteger)
 tdf.Fields.Append fldG11
 '  Call SetPropertyDAO(fldG11, "DisplayControl", dbInteger, CInt(acCheckBox))
 CaLL yesNoNoneQuestions(fldG11)

 Dim fldG12 As DAO.Field
 Set fldG12 = tdf.CreateField("MH_11", dbInteger)
 tdf.Fields.Append fldG12
 '  Call SetPropertyDAO(fldG12, "DisplayControl", dbInteger, CInt(acCheckBox))
 CaLL yesNoNoneQuestions(fldG12)

 Dim fldG13 As DAO.Field
 Set fldG13 = tdf.CreateField("MH_12", dbInteger)
 tdf.Fields.Append fldG13
 '  Call SetPropertyDAO(fldG13, "DisplayControl", dbInteger, CInt(acCheckBox))
 CaLL yesNoNoneQuestions(fldG13)

 Dim fldG14 As DAO.Field
 Set fldG14 = tdf.CreateField("MH_13", dbInteger)
 tdf.Fields.Append fldG14
 '  Call SetPropertyDAO(fldG14, "DisplayControl", dbInteger, CInt(acCheckBox))
 CaLL yesNoNoneQuestions(fldG14)

 Dim fldG15 As DAO.Field
 Set fldG15 = tdf.CreateField("MH_14", dbInteger)
 tdf.Fields.Append fldG15
 '  Call SetPropertyDAO(fldG15, "DisplayControl", dbInteger, CInt(acCheckBox))
 CaLL yesNoNoneQuestions(fldG15)

 Dim fldG16 As DAO.Field
 Set fldG16 = tdf.CreateField("MH_15", dbInteger)
 tdf.Fields.Append fldG16
 '  Call SetPropertyDAO(fldG16, "DisplayControl", dbInteger, CInt(acCheckBox))
 CaLL yesNoNoneQuestions(fldG16)

 Dim fldG17 As DAO.Field
 Set fldG17 = tdf.CreateField("MH_16", dbInteger)
 tdf.Fields.Append fldG17
 '  Call SetPropertyDAO(fldG17, "DisplayControl", dbInteger, CInt(acCheckBox))
 CaLL yesNoNoneQuestions(fldG17)

 Dim fldG18 As DAO.Field
 Set fldG18 = tdf.CreateField("MH_17", dbInteger)
 tdf.Fields.Append fldG18
 '  Call SetPropertyDAO(fldG18, "DisplayControl", dbInteger, CInt(acCheckBox))
 CaLL yesNoNoneQuestions(fldG18)
 
 'SUD screen fields
 Dim fldH01 As DAO.Field
 Set fldH01 = tdf.CreateField("SUD_1", dbInteger)
 tdf.Fields.Append fldH01
 '  Call SetPropertyDAO(fldH01, "DisplayControl", dbInteger, CInt(acCheckBox))
Call yesNoNoneQuestions(fldH01)

 Dim fldH02 As DAO.Field
 Set fldH02 = tdf.CreateField("SUD_2", dbInteger)
 tdf.Fields.Append fldH02
 '  Call SetPropertyDAO(fldH02, "DisplayControl", dbInteger, CInt(acCheckBox))
Call yesNoNoneQuestions(fldH02)

 Dim fldH03 As DAO.Field
 Set fldH03 = tdf.CreateField("SUD_3", dbInteger)
 tdf.Fields.Append fldH03
 '  Call SetPropertyDAO(fldH03, "DisplayControl", dbInteger, CInt(acCheckBox))
Call yesNoNoneQuestions(fldH03)

 Dim fldH04 As DAO.Field
 Set fldH04 = tdf.CreateField("SUD_4", dbInteger)
 tdf.Fields.Append fldH04
 '  Call SetPropertyDAO(fldH04, "DisplayControl", dbInteger, CInt(acCheckBox))
Call yesNoNoneQuestions(fldH04)

 Dim fldH05 As DAO.Field
 Set fldH05 = tdf.CreateField("SUD_5a", dbBoolean)
 tdf.Fields.Append fldH05
 Call SetPropertyDAO(fldH05, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldH06 As DAO.Field
 Set fldH06 = tdf.CreateField("SUD_5b", dbBoolean)
 tdf.Fields.Append fldH06
 Call SetPropertyDAO(fldH06, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldH07 As DAO.Field
 Set fldH07 = tdf.CreateField("SUD_5c", dbBoolean)
 tdf.Fields.Append fldH07
 Call SetPropertyDAO(fldH07, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldH08 As DAO.Field
 Set fldH08 = tdf.CreateField("SUD_5d", dbBoolean)
 tdf.Fields.Append fldH08
 Call SetPropertyDAO(fldH08, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldH09 As DAO.Field
 Set fldH09 = tdf.CreateField("SUD_5e", dbBoolean)
 tdf.Fields.Append fldH09
 Call SetPropertyDAO(fldH09, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldH10 As DAO.Field
 Set fldH10 = tdf.CreateField("SUD_5f", dbBoolean)
 tdf.Fields.Append fldH10
 Call SetPropertyDAO(fldH10, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldH11 As DAO.Field
 Set fldH11 = tdf.CreateField("SUD_5g", dbBoolean)
 tdf.Fields.Append fldH11
 Call SetPropertyDAO(fldH11, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldH12 As DAO.Field
 Set fldH12 = tdf.CreateField("SUD_5h", dbBoolean)
 tdf.Fields.Append fldH12
 Call SetPropertyDAO(fldH12, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldH13 As DAO.Field
 Set fldH13 = tdf.CreateField("SUD_5", dbInteger)
 tdf.Fields.Append fldH13
 '  Call SetPropertyDAO(fldH13, "DisplayControl", dbInteger, CInt(acCheckBox))
 Call yesNoNoneQuestions(fldH13)

 Dim fldH14 As DAO.Field
 Set fldH14 = tdf.CreateField("SUD_6", dbInteger)
 tdf.Fields.Append fldH14
 '  Call SetPropertyDAO(fldH14, "DisplayControl", dbInteger, CInt(acCheckBox))
 Call yesNoNoneQuestions(fldH14)

 Dim fldH15 As DAO.Field
 Set fldH15 = tdf.CreateField("SUD_7", dbInteger)
 tdf.Fields.Append fldH15
 '  Call SetPropertyDAO(fldH15, "DisplayControl", dbInteger, CInt(acCheckBox))
 Call yesNoNoneQuestions(fldH15)

 Dim fldH16 As DAO.Field
 Set fldH16 = tdf.CreateField("SUD_8", dbInteger)
 tdf.Fields.Append fldH16
 '  Call SetPropertyDAO(fldH16, "DisplayControl", dbInteger, CInt(acCheckBox))
 Call yesNoNoneQuestions(fldH16)

 Dim fldH17 As DAO.Field
 Set fldH17 = tdf.CreateField("SUD_9", dbInteger)
 tdf.Fields.Append fldH17
 '  Call SetPropertyDAO(fldH17, "DisplayControl", dbInteger, CInt(acCheckBox))
 Call yesNoNoneQuestions(fldH17)

 Dim fldH18 As DAO.Field
 Set fldH18 = tdf.CreateField("SUD_10", dbInteger)
 tdf.Fields.Append fldH18
 '  Call SetPropertyDAO(fldH18, "DisplayControl", dbInteger, CInt(acCheckBox))
 Call yesNoNoneQuestions(fldH18)

 Dim fldH19 As DAO.Field
 Set fldH19 = tdf.CreateField("SUD_11", dbInteger)
 tdf.Fields.Append fldH19
 '  Call SetPropertyDAO(fldH19, "DisplayControl", dbInteger, CInt(acCheckBox))
 Call yesNoNoneQuestions(fldH19)

 Dim fldH20 As DAO.Field
 Set fldH20 = tdf.CreateField("SUD_12", dbInteger)
 tdf.Fields.Append fldH20
 '  Call SetPropertyDAO(fldH20, "DisplayControl", dbInteger, CInt(acCheckBox))
 Call yesNoNoneQuestions(fldH20)

 Dim fldH21 As DAO.Field
 Set fldH21 = tdf.CreateField("SUD_13", dbInteger)
 tdf.Fields.Append fldH21
 '  Call SetPropertyDAO(fldH21, "DisplayControl", dbInteger, CInt(acCheckBox))
 Call yesNoNoneQuestions(fldH21)

 Dim fldH22 As DAO.Field
 Set fldH22 = tdf.CreateField("SUD_14", dbInteger)
 tdf.Fields.Append fldH22
 '  Call SetPropertyDAO(fldH22, "DisplayControl", dbInteger, CInt(acCheckBox))
 Call yesNoNoneQuestions(fldH22)

 Dim fldH23 As DAO.Field
 Set fldH23 = tdf.CreateField("SUD_15", dbInteger)
 tdf.Fields.Append fldH23
 '  Call SetPropertyDAO(fldH23, "DisplayControl", dbInteger, CInt(acCheckBox))
 Call yesNoNoneQuestions(fldH23)

 Dim fldH24 As DAO.Field
 Set fldH24 = tdf.CreateField("SUD_16", dbInteger)
 tdf.Fields.Append fldH24
 '  Call SetPropertyDAO(fldH24, "DisplayControl", dbInteger, CInt(acCheckBox))
 Call yesNoNoneQuestions(fldH24)

 Dim fldH25 As DAO.Field
 Set fldH25 = tdf.CreateField("Prim_Sub", dbText)
 tdf.Fields.Append fldH25
 Call setComboProperties(fldH25, "Not collected;None;Alcohol;Amphetamines;Barbiturates;Benzodiazepines;Cocaine/crack;Hallucinogens;Heroin;Inhalants;Marijuana;Methamphetamines;Non-prescriptive methadone;Other amphetamines;Other hallucinogens;Others sedatives;Other tranquilizers;Other opiates;Other stimulants;OTC drugs;PCP;Other")

 Dim fldH26 As DAO.Field
 Set fldH26 = tdf.CreateField("Prim_age_first", dbInteger)
 tdf.Fields.Append fldH26

 Dim fldH27 As DAO.Field
 Set fldH27 = tdf.CreateField("Prim_Sev", dbText)
 tdf.Fields.Append fldH27
 Call setComboProperties(fldH27, "Severe;Moderate;Mild;Not a problem;NA;Not collected")

 Dim fldH28 As DAO.Field
 Set fldH28 = tdf.CreateField("Prim_Freq", dbText)
 tdf.Fields.Append fldH28
 Call setComboProperties(fldH28, "None in past month;1-3 x in past month;1-2 x in past week;3-6 x in past week;Daily;NA;Unknown;Not collected")

 Dim fldH29 As DAO.Field
 Set fldH29 = tdf.CreateField("Prim_Meth", dbText)
 tdf.Fields.Append fldH29
 Call setComboProperties(fldH29, "Oral;Smoking;Inhalation;Injection;Other;NA;Unknown;Not collected")

 Dim fldH30 As DAO.Field
 Set fldH30 = tdf.CreateField("Prim_days_month", dbInteger)
 tdf.Fields.Append fldH30

 Dim fldH31 As DAO.Field
 Set fldH31 = tdf.CreateField("Prim_last_use_month", dbInteger)
 tdf.Fields.Append fldH31

 Dim fldH31a As DAO.Field
 Set fldH31a = tdf.CreateField("Prim_last_use_year", dbInteger)
 tdf.Fields.Append fldH31a

 Dim fldH32 As DAO.Field
 Set fldH32 = tdf.CreateField("Sec_Sub", dbText)
 tdf.Fields.Append fldH32
 Call setComboProperties(fldH32, "N/A;Not collected;None;Alcohol;Amphetamines;Barbiturates;Benzodiazepines;Cocaine/crack;Hallucinogens;Heroin;Inhalants;Marijuana;Methamphetamines;Non-prescriptive methadone;Other amphetamines;Other hallucinogens;Others sedatives;Other tranquilizers;Other opiates;Other stimulants;OTC drugs;PCP;Other")

 Dim fldH33 As DAO.Field
 Set fldH33 = tdf.CreateField("Sec_age_first", dbInteger)
 tdf.Fields.Append fldH33

 Dim fldH34 As DAO.Field
 Set fldH34 = tdf.CreateField("Sec_Sev", dbText)
 tdf.Fields.Append fldH34
 Call setComboProperties(fldH34, "Severe;Moderate;Mild;Not a problem;NA;Not collected")

 Dim fldH35 As DAO.Field
 Set fldH35 = tdf.CreateField("Sec_Freq", dbText)
 tdf.Fields.Append fldH35
 Call setComboProperties(fldH35, "None in past month;1-3 x in past month;1-2 x in past week;3-6 x in past week;Daily;NA;Unknown;Not collected")

 Dim fldH36 As DAO.Field
 Set fldH36 = tdf.CreateField("Sec_Meth", dbText)
 tdf.Fields.Append fldH36
 Call setComboProperties(fldH36, "Oral;Smoking;Inhalation;Injection;Other;NA;Unknown;Not collected")

 Dim fldH37 As DAO.Field
 Set fldH37 = tdf.CreateField("Sec_days_month", dbInteger)
 tdf.Fields.Append fldH37

 Dim fldH38 As DAO.Field
 Set fldH38 = tdf.CreateField("Sec_last_use_month", dbInteger)
 tdf.Fields.Append fldH38

 Dim fldH38a As DAO.Field
 Set fldH38a = tdf.CreateField("Sec_last_use_year", dbDdbIntegerate)
 tdf.Fields.Append fldH38a

 Dim fldH39 As DAO.Field
 Set fldH39 = tdf.CreateField("Other_Addictions_Food", dbBoolean)
 tdf.Fields.Append fldH39
 Call SetPropertyDAO(fldH39, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldH40 As DAO.Field
 Set fldH40 = tdf.CreateField("Other_Addictions_Gambling", dbBoolean)
 tdf.Fields.Append fldH40
 Call SetPropertyDAO(fldH40, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldH41 As DAO.Field
 Set fldH41 = tdf.CreateField("Other_Addictions_Prescription", dbBoolean)
 tdf.Fields.Append fldH41
 Call SetPropertyDAO(fldH41, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldH42 As DAO.Field
 Set fldH42 = tdf.CreateField("Other_Addictions_Street", dbBoolean)
 tdf.Fields.Append fldH42
 Call SetPropertyDAO(fldH42, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldH43 As DAO.Field
 Set fldH43 = tdf.CreateField("LT_out_admin", dbInteger)
 tdf.Fields.Append fldH43

 Dim fldH44 As DAO.Field
 Set fldH44 = tdf.CreateField("LT_in_admin", dbInteger)
 tdf.Fields.Append fldH44

 Dim fldH45 As DAO.Field
 Set fldH45 = tdf.CreateField("Needle_track_marks", dbInteger)
 tdf.Fields.Append fldH45
 '  Call SetPropertyDAO(fldH45, "DisplayControl", dbInteger, CInt(acCheckBox))
 Call yesNoNoneQuestions(fldH45)

 Dim fldH46 As DAO.Field
 Set fldH46 = tdf.CreateField("Skin_abscesses_etc", dbInteger)
 tdf.Fields.Append fldH46
 '  Call SetPropertyDAO(fldH46, "DisplayControl", dbInteger, CInt(acCheckBox))
 Call yesNoNoneQuestions(fldH46)

 Dim fldH47 As DAO.Field
 Set fldH47 = tdf.CreateField("Tremors", dbInteger)
 tdf.Fields.Append fldH47
 '  Call SetPropertyDAO(fldH47, "DisplayControl", dbInteger, CInt(acCheckBox))
 Call yesNoNoneQuestions(fldH47)

 Dim fldH48 As DAO.Field
 Set fldH48 = tdf.CreateField("Unclear_speech", dbInteger)
 tdf.Fields.Append fldH48
 '  Call SetPropertyDAO(fldH48, "DisplayControl", dbInteger, CInt(acCheckBox))
 Call yesNoNoneQuestions(fldH48)

 Dim fldH49 As DAO.Field
 Set fldH49 = tdf.CreateField("Unsteady_gait", dbInteger)
 tdf.Fields.Append fldH49
 '  Call SetPropertyDAO(fldH49, "DisplayControl", dbInteger, CInt(acCheckBox))
 Call yesNoNoneQuestions(fldH49)

 Dim fldH50 As DAO.Field
 Set fldH50 = tdf.CreateField("Dialated_Pupils", dbInteger)
 tdf.Fields.Append fldH50
 '  Call SetPropertyDAO(fldH50, "DisplayControl", dbInteger, CInt(acCheckBox))
 Call yesNoNoneQuestions(fldH50)

 Dim fldH51 As DAO.Field
 Set fldH51 = tdf.CreateField("Scratching", dbInteger)
 tdf.Fields.Append fldH51
 '  Call SetPropertyDAO(fldH51, "DisplayControl", dbInteger, CInt(acCheckBox))
 Call yesNoNoneQuestions(fldH51)

 Dim fldH52 As DAO.Field
 Set fldH52 = tdf.CreateField("Swollen_hands_or_feet", dbInteger)
 tdf.Fields.Append fldH52
 '  Call SetPropertyDAO(fldH52, "DisplayControl", dbInteger, CInt(acCheckBox))
 Call yesNoNoneQuestions(fldH52)

 Dim fldH53 As DAO.Field
 Set fldH53 = tdf.CreateField("Smell_of_alcohol_or_marijuana", dbInteger)
 tdf.Fields.Append fldH53
 '  Call SetPropertyDAO(fldH53, "DisplayControl", dbInteger, CInt(acCheckBox))
 Call yesNoNoneQuestions(fldH53)

 Dim fldH54 As DAO.Field
 Set fldH54 = tdf.CreateField("Drug_paraphernalia", dbInteger)
 tdf.Fields.Append fldH54
 '  Call SetPropertyDAO(fldH54, "DisplayControl", dbInteger, CInt(acCheckBox))
 Call yesNoNoneQuestions(fldH54)

 Dim fldH55 As DAO.Field
 Set fldH55 = tdf.CreateField("Nodding_out", dbInteger)
 tdf.Fields.Append fldH55
 '  Call SetPropertyDAO(fldH55, "DisplayControl", dbInteger, CInt(acCheckBox))
 Call yesNoNoneQuestions(fldH55)

 Dim fldH56 As DAO.Field
 Set fldH56 = tdf.CreateField("Agitation", dbInteger)
 tdf.Fields.Append fldH56
 '  Call SetPropertyDAO(fldH56, "DisplayControl", dbInteger, CInt(acCheckBox))
 Call yesNoNoneQuestions(fldH56)

 Dim fldH57 As DAO.Field
 Set fldH57 = tdf.CreateField("Inability_to_focus", dbInteger)
 tdf.Fields.Append fldH57
 '  Call SetPropertyDAO(fldH57, "DisplayControl", dbInteger, CInt(acCheckBox))
 Call yesNoNoneQuestions(fldH57)

 Dim fldH58 As DAO.Field
 Set fldH58 = tdf.CreateField("Burns_on_inside_of_lips", dbInteger)
 tdf.Fields.Append fldH58
 '  Call SetPropertyDAO(fldH58, "DisplayControl", dbInteger, CInt(acCheckBox))
 Call yesNoNoneQuestions(fldH58)

 '---------------------------------------------------------------
 '---------------------------------------------------------------
 '---------------------------------------------------------------
 '---------------------------------------------------------------
 '---------------------------------------------------------------
 '---------------------------------------------------------------

 'create the table 2 definition
 Set tdf = db.CreateTableDef("SATU_db_2")
 
 'create the field definitions
 Dim fldID2 As DAO.Field
 Set fldID2 = tdf.CreateField("autoID", dbLong)
 fldID2.Attributes = dbAutoIncrField
 fldID2.Required = True
 tdf.Fields.Append fldID2
 
 'add the table to the database
 db.TableDefs.Append tdf
 Set tdf = db.TableDefs("SATU_db_2")
 
 'add other fields
 Dim fldB01a As DAO.Field
 Set fldB01a = tdf.CreateField("ResearchID_2", dbText)
 tdf.Fields.Append fldB01a

 'CAF I fields
 Dim fldI001 As DAO.Field
 Set fldI001 = tdf.CreateField("Date_of_visit", dbDate)
 tdf.Fields.Append fldI001

 Dim fldI002 As DAO.Field
 Set fldI002 = tdf.CreateField("Start", dbDate)
 tdf.Fields.Append fldI002

 Dim fldI003 As DAO.Field
 Set fldI003 = tdf.CreateField("End", dbDate)
 tdf.Fields.Append fldI003

 Dim fldI004 As DAO.Field
 Set fldI004 = tdf.CreateField("Clinician", dbText)
 tdf.Fields.Append fldI004

 Dim fldI005 As DAO.Field
 Set fldI005 = tdf.CreateField("ICD-Mental_health", dbText)
 tdf.Fields.Append fldI005
 '  Call setComboMultiProperties(fldI005, 0)

 Dim fldI006 As DAO.Field
 Set fldI006 = tdf.CreateField("Chief_complaint", dbMemo)
 tdf.Fields.Append fldI006

 Dim fldI007 As DAO.Field
 Set fldI007 = tdf.CreateField("Current_legal", dbBoolean)
 tdf.Fields.Append fldI007
 Call SetPropertyDAO(fldI007, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldI008 As DAO.Field
 Set fldI008 = tdf.CreateField("C-SSRS_1", dbBoolean)
 tdf.Fields.Append fldI008
 Call SetPropertyDAO(fldI008, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldI009 As DAO.Field
 Set fldI009 = tdf.CreateField("C-SSRS_2", dbBoolean)
 tdf.Fields.Append fldI009
 Call SetPropertyDAO(fldI009, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldI010 As DAO.Field
 Set fldI010 = tdf.CreateField("C-SSRS_3", dbBoolean)
 tdf.Fields.Append fldI010
 Call SetPropertyDAO(fldI010, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldI011 As DAO.Field
 Set fldI011 = tdf.CreateField("C-SSRS_4", dbBoolean)
 tdf.Fields.Append fldI011
 Call SetPropertyDAO(fldI011, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldI012 As DAO.Field
 Set fldI012 = tdf.CreateField("C_SSRS_5", dbBoolean)
 tdf.Fields.Append fldI012
 Call SetPropertyDAO(fldI012, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldI013 As DAO.Field
 Set fldI013 = tdf.CreateField("C_SSRS_Sui_Bx_LT", dbBoolean)
 tdf.Fields.Append fldI013
 Call SetPropertyDAO(fldI013, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldI014 As DAO.Field
 Set fldI014 = tdf.CreateField("C-SSRS_Siu_Bx_Past_3_Months", dbBoolean)
 tdf.Fields.Append fldI014
 Call SetPropertyDAO(fldI014, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldI015 As DAO.Field
 Set fldI015 = tdf.CreateField("RF_Psych_mood", dbBoolean)
 tdf.Fields.Append fldI015
 Call SetPropertyDAO(fldI015, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldI016 As DAO.Field
 Set fldI016 = tdf.CreateField("RF_Psych_psychotic", dbBoolean)
 tdf.Fields.Append fldI016
 Call SetPropertyDAO(fldI016, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldI017 As DAO.Field
 Set fldI017 = tdf.CreateField("RF_Psych_sud", dbBoolean)
 tdf.Fields.Append fldI017
 Call SetPropertyDAO(fldI017, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldI018 As DAO.Field
 Set fldI018 = tdf.CreateField("RF_Psych_ptsd", dbBoolean)
 tdf.Fields.Append fldI018
 Call SetPropertyDAO(fldI018, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldI019 As DAO.Field
 Set fldI019 = tdf.CreateField("RF_Psych_adhd", dbBoolean)
 tdf.Fields.Append fldI019
 Call SetPropertyDAO(fldI019, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldI020 As DAO.Field
 Set fldI020 = tdf.CreateField("RF_Psych_tbi", dbBoolean)
 tdf.Fields.Append fldI020
 Call SetPropertyDAO(fldI020, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldI021 As DAO.Field
 Set fldI021 = tdf.CreateField("RF_Psych_clusB", dbBoolean)
 tdf.Fields.Append fldI021
 Call SetPropertyDAO(fldI021, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldI022 As DAO.Field
 Set fldI022 = tdf.CreateField("RF_Psych_cond", dbBoolean)
 tdf.Fields.Append fldI022
 Call SetPropertyDAO(fldI022, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldI023 As DAO.Field
 Set fldI023 = tdf.CreateField("RF_Psych_rec-on", dbBoolean)
 tdf.Fields.Append fldI023
 Call SetPropertyDAO(fldI023, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldI024 As DAO.Field
 Set fldI024 = tdf.CreateField("RF_Sxs_anhed", dbBoolean)
 tdf.Fields.Append fldI024
 Call SetPropertyDAO(fldI024, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldI025 As DAO.Field
 Set fldI025 = tdf.CreateField("RF_Sxs_impul", dbBoolean)
 tdf.Fields.Append fldI025
 Call SetPropertyDAO(fldI025, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldI026 As DAO.Field
 Set fldI026 = tdf.CreateField("RF_Sxs_hopel", dbBoolean)
 tdf.Fields.Append fldI026
 Call SetPropertyDAO(fldI026, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldI027 As DAO.Field
 Set fldI027 = tdf.CreateField("RF_Sxs_anx", dbBoolean)
 tdf.Fields.Append fldI027
 Call SetPropertyDAO(fldI027, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldI028 As DAO.Field
 Set fldI028 = tdf.CreateField("RF_Sxs_insom", dbBoolean)
 tdf.Fields.Append fldI028
 Call SetPropertyDAO(fldI028, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldI029 As DAO.Field
 Set fldI029 = tdf.CreateField("RF_Sxs_halluc", dbBoolean)
 tdf.Fields.Append fldI029
 Call SetPropertyDAO(fldI029, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldI030 As DAO.Field
 Set fldI030 = tdf.CreateField("RF_Sxs_psychosis", dbBoolean)
 tdf.Fields.Append fldI030
 Call SetPropertyDAO(fldI030, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldI031 As DAO.Field
 Set fldI031 = tdf.CreateField("RF_Sxs_akathi", dbBoolean)
 tdf.Fields.Append fldI031
 Call SetPropertyDAO(fldI031, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldI032 As DAO.Field
 Set fldI032 = tdf.CreateField("RF_Sxs_other", dbBoolean)
 tdf.Fields.Append fldI032
 Call SetPropertyDAO(fldI032, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldI033 As DAO.Field
 Set fldI033 = tdf.CreateField("RF_FamHx_suic", dbBoolean)
 tdf.Fields.Append fldI033
 Call SetPropertyDAO(fldI033, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldI034 As DAO.Field
 Set fldI034 = tdf.CreateField("RF_FamHx_suic-beh", dbBoolean)
 tdf.Fields.Append fldI034
 Call SetPropertyDAO(fldI034, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldI035 As DAO.Field
 Set fldI035 = tdf.CreateField("RF_FamHx_psy-diag", dbBoolean)
 tdf.Fields.Append fldI035
 Call SetPropertyDAO(fldI035, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldI036 As DAO.Field
 Set fldI036 = tdf.CreateField("RF_Stress_trigg", dbBoolean)
 tdf.Fields.Append fldI036
 Call SetPropertyDAO(fldI036, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldI037 As DAO.Field
 Set fldI037 = tdf.CreateField("RF_Stress_pain", dbBoolean)
 tdf.Fields.Append fldI037
 Call SetPropertyDAO(fldI037, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldI038 As DAO.Field
 Set fldI038 = tdf.CreateField("RF_Stress_sex-abu", dbBoolean)
 tdf.Fields.Append fldI038
 Call SetPropertyDAO(fldI038, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldI039 As DAO.Field
 Set fldI039 = tdf.CreateField("RF_Stress_intox", dbBoolean)
 tdf.Fields.Append fldI039
 Call SetPropertyDAO(fldI039, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldI040 As DAO.Field
 Set fldI040 = tdf.CreateField("RF_Stress_incarc-home", dbBoolean)
 tdf.Fields.Append fldI040
 Call SetPropertyDAO(fldI040, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldI041 As DAO.Field
 Set fldI041 = tdf.CreateField("RF_Stress_legal", dbBoolean)
 tdf.Fields.Append fldI041
 Call SetPropertyDAO(fldI041, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldI042 As DAO.Field
 Set fldI042 = tdf.CreateField("RF_Stress_soc-sup", dbBoolean)
 tdf.Fields.Append fldI042
 Call SetPropertyDAO(fldI042, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldI043 As DAO.Field
 Set fldI043 = tdf.CreateField("RF_Stress_soc-iso", dbBoolean)
 tdf.Fields.Append fldI043
 Call SetPropertyDAO(fldI043, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldI044 As DAO.Field
 Set fldI044 = tdf.CreateField("RF_Stress_burden", dbBoolean)
 tdf.Fields.Append fldI044
 Call SetPropertyDAO(fldI044, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldI045 As DAO.Field
 Set fldI045 = tdf.CreateField("RF_tx_inp-disc", dbBoolean)
 tdf.Fields.Append fldI045
 Call SetPropertyDAO(fldI045, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldI046 As DAO.Field
 Set fldI046 = tdf.CreateField("RF_tx_chng-treat", dbBoolean)
 tdf.Fields.Append fldI046
 Call SetPropertyDAO(fldI046, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldI047 As DAO.Field
 Set fldI047 = tdf.CreateField("RF_tx_diss-treat", dbBoolean)
 tdf.Fields.Append fldI047
 Call SetPropertyDAO(fldI047, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldI048 As DAO.Field
 Set fldI048 = tdf.CreateField("RF_tx_no-treat", dbBoolean)
 tdf.Fields.Append fldI048
 Call SetPropertyDAO(fldI048, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldI049 As DAO.Field
 Set fldI049 = tdf.CreateField("RF_lethal", dbBoolean)
 tdf.Fields.Append fldI049
 Call SetPropertyDAO(fldI049, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldI050 As DAO.Field
 Set fldI050 = tdf.CreateField("PF_Int_cope", dbBoolean)
 tdf.Fields.Append fldI050
 Call SetPropertyDAO(fldI050, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldI051 As DAO.Field
 Set fldI051 = tdf.CreateField("PF_Int_frust", dbBoolean)
 tdf.Fields.Append fldI051
 Call SetPropertyDAO(fldI051, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldI052 As DAO.Field
 Set fldI052 = tdf.CreateField("PF_Int_relig", dbBoolean)
 tdf.Fields.Append fldI052
 Call SetPropertyDAO(fldI052, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldI053 As DAO.Field
 Set fldI053 = tdf.CreateField("PF_Int_fear", dbBoolean)
 tdf.Fields.Append fldI053
 Call SetPropertyDAO(fldI053, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldI054 As DAO.Field
 Set fldI054 = tdf.CreateField("PF_Int_reason", dbBoolean)
 tdf.Fields.Append fldI054
 Call SetPropertyDAO(fldI054, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldI055 As DAO.Field
 Set fldI055 = tdf.CreateField("PF_Ext_cultur", dbBoolean)
 tdf.Fields.Append fldI055
 Call SetPropertyDAO(fldI055, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldI056 As DAO.Field
 Set fldI056 = tdf.CreateField("PF_Ext_respo", dbBoolean)
 tdf.Fields.Append fldI056
 Call SetPropertyDAO(fldI056, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldI057 As DAO.Field
 Set fldI057 = tdf.CreateField("PF_Ext_pets", dbBoolean)
 tdf.Fields.Append fldI057
 Call SetPropertyDAO(fldI057, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldI058 As DAO.Field
 Set fldI058 = tdf.CreateField("PF_Ext_soc-net", dbBoolean)
 tdf.Fields.Append fldI058
 Call SetPropertyDAO(fldI058, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldI059 As DAO.Field
 Set fldI059 = tdf.CreateField("PF_Ext_thera-rel", dbBoolean)
 tdf.Fields.Append fldI059
 Call SetPropertyDAO(fldI059, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldI060 As DAO.Field
 Set fldI060 = tdf.CreateField("PF_Ext_work", dbBoolean)
 tdf.Fields.Append fldI060
 Call SetPropertyDAO(fldI060, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldI061 As DAO.Field
 Set fldI061 = tdf.CreateField("C_SSRS_Step3_1", dbText)
 tdf.Fields.Append fldI061
 Call setComboProperties(fldI061, "Less than once a week;Once a week;2-5 times in week;Daily or almost daily;Many times a day")

 Dim fldI062 As DAO.Field
 Set fldI062 = tdf.CreateField("C_SSRS_Step3_2", dbText)
 tdf.Fields.Append fldI062
 Call setComboProperties(fldI062, "Fleeting;Less than one hour;1-4 hours;4-8 hours;More than 8 hours")

 Dim fldI063 As DAO.Field
 Set fldI063 = tdf.CreateField("C_SSRS_Step3_3", dbText)
 tdf.Fields.Append fldI063
 Call setComboProperties(fldI063, "Easily able;Little difficulty;Some difficulty;Lot of difficulty;Unable;Does not attempt")

 Dim fldI064 As DAO.Field
 Set fldI064 = tdf.CreateField("C_SSRS_Step3_4", dbText)
 tdf.Fields.Append fldI064
 Call setComboProperties(fldI064, "Definitely stopped;Probably stopped;Uncertain;Most likely did not;Definitely did not;Does not apply")

 Dim fldI065 As DAO.Field
 Set fldI065 = tdf.CreateField("C_SSRS_Step3_5", dbText)
 tdf.Fields.Append fldI065
 Call setComboProperties(fldI065, "Completely to get attention;Mostly to get attention;Equally to get attention/end;Mostly to end;Completely to end;Does not apply")

 Dim fldI066 As DAO.Field
 Set fldI066 = tdf.CreateField("Risk_stratification", dbText)
 tdf.Fields.Append fldI066
 Call setComboProperties(fldI066, "HIGH-past month;HIGH-past 3 months;MOD-without plan; MOD->3 months;MOD-multiple and few;LOW-without method;LOW-modifiable and strong;LOW-no history")

 Dim fldI067 As DAO.Field
 Set fldI067 = tdf.CreateField("History_of_violence", dbBoolean)
 tdf.Fields.Append fldI067
 Call SetPropertyDAO(fldI067, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldI068 As DAO.Field
 Set fldI068 = tdf.CreateField("History_fire_setting", dbBoolean)
 tdf.Fields.Append fldI068
 Call SetPropertyDAO(fldI068, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldI069 As DAO.Field
 Set fldI069 = tdf.CreateField("Others_complained", dbBoolean)
 tdf.Fields.Append fldI069
 Call SetPropertyDAO(fldI069, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldI070 As DAO.Field
 Set fldI070 = tdf.CreateField("Harm_in_past", dbBoolean)
 tdf.Fields.Append fldI070
 Call SetPropertyDAO(fldI070, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldI071 As DAO.Field
 Set fldI071 = tdf.CreateField("Current_HI", dbBoolean)
 tdf.Fields.Append fldI071
 Call SetPropertyDAO(fldI071, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldI072 As DAO.Field
 Set fldI072 = tdf.CreateField("Current_homicidal_plan", dbBoolean)
 tdf.Fields.Append fldI072
 Call SetPropertyDAO(fldI072, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldI073 As DAO.Field
 Set fldI073 = tdf.CreateField("traum_sxs1", dbInteger)
 tdf.Fields.Append fldI073
 '  Call SetPropertyDAO(fldI073, "DisplayControl", dbInteger, CInt(acCheckBox))
 Call yesNoNoneQuestions(fldI073)

 Dim fldI074 As DAO.Field
 Set fldI074 = tdf.CreateField("traum_sxs2", dbInteger)
 tdf.Fields.Append fldI074
 '  Call SetPropertyDAO(fldI074, "DisplayControl", dbInteger, CInt(acCheckBox))
 Call yesNoNoneQuestions(fldI074)

 Dim fldI075 As DAO.Field
 Set fldI075 = tdf.CreateField("traum_sxs3", dbInteger)
 tdf.Fields.Append fldI075
 '  Call SetPropertyDAO(fldI075, "DisplayControl", dbInteger, CInt(acCheckBox))
 Call yesNoNoneQuestions(fldI075)

 Dim fldI076 As DAO.Field
 Set fldI076 = tdf.CreateField("traum_sxs4", dbInteger)
 tdf.Fields.Append fldI076
 '  Call SetPropertyDAO(fldI076, "DisplayControl", dbInteger, CInt(acCheckBox))
 Call yesNoNoneQuestions(fldI076)

 Dim fldI077 As DAO.Field
 Set fldI077 = tdf.CreateField("traum_sxs5", dbInteger)
 tdf.Fields.Append fldI077
 '  Call SetPropertyDAO(fldI077, "DisplayControl", dbInteger, CInt(acCheckBox))
 Call yesNoNoneQuestions(fldI077)

 Dim fldI078 As DAO.Field
 Set fldI078 = tdf.CreateField("traum_sxs6", dbInteger)
 tdf.Fields.Append fldI078
 '  Call SetPropertyDAO(fldI078, "DisplayControl", dbInteger, CInt(acCheckBox))
 Call yesNoNoneQuestions(fldI078)

 Dim fldI079 As DAO.Field
 Set fldI079 = tdf.CreateField("Medical_conditions", dbText)
 tdf.Fields.Append fldI079
 Call setComboMultiProperties(fldI079, "None;Arthritis;Asthma;Cancer;Diabetes;GERD;Head trauma;Heart disease;Hepatitis A;Hepatitis B;Hepatitis C;High blood pressure;HIV;HX + PPD;Liver disease;Lung disease;Obesity;Osteoporosis;Pneumonia;Renal problems;Recent/major medical/surgical;seizures;Sleep apnea;Thyroid disease;Tuberculosis;Other")

 Dim fldI080 As DAO.Field
 Set fldI080 = tdf.CreateField("Vis_Hear_Imp", dbBoolean)
 tdf.Fields.Append fldI080
 Call SetPropertyDAO(fldI080, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldI081 As DAO.Field
 Set fldI081 = tdf.CreateField("PrimaryCare", dbBoolean)
 tdf.Fields.Append fldI081
 Call SetPropertyDAO(fldI081, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldI082 As DAO.Field
 Set fldI082 = tdf.CreateField("Date_physical", dbDate)
 tdf.Fields.Append fldI082

 Dim fldI083 As DAO.Field
 Set fldI083 = tdf.CreateField("Date_dental", dbDate)
 tdf.Fields.Append fldI083

 Dim fldI084 As DAO.Field
 Set fldI084 = tdf.CreateField("Sex_active", dbBoolean)
 tdf.Fields.Append fldI084
 Call SetPropertyDAO(fldI084, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldI085 As DAO.Field
 Set fldI085 = tdf.CreateField("Safe_sex", dbBoolean)
 tdf.Fields.Append fldI085
 Call SetPropertyDAO(fldI085, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldI086 As DAO.Field
 Set fldI086 = tdf.CreateField("date_menses", dbDate)
 tdf.Fields.Append fldI086

 Dim fldI087 As DAO.Field
 Set fldI087 = tdf.CreateField("withdrawal", dbBoolean)
 tdf.Fields.Append fldI087
 Call SetPropertyDAO(fldI087, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldI088 As DAO.Field
 Set fldI088 = tdf.CreateField("Pain_daily", dbBoolean)
 tdf.Fields.Append fldI088
 Call SetPropertyDAO(fldI088, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldI089 As DAO.Field
 Set fldI089 = tdf.CreateField("Pain_cur", dbBoolean)
 tdf.Fields.Append fldI089
 Call SetPropertyDAO(fldI089, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldI090 As DAO.Field
 Set fldI090 = tdf.CreateField("Pain_cur_rate", dbText)
 tdf.Fields.Append fldI090
 Call setComboProperties(fldI090, "No pain;Mild pain;Mild/moderate pain;Moderate pain;Moderate/severe pain;Severe pain")

 Dim fldI091 As DAO.Field
 Set fldI091 = tdf.CreateField("height_caf", dbText)
 tdf.Fields.Append fldI091

 Dim fldI092 As DAO.Field
 Set fldI092 = tdf.CreateField("weight_caf", dbInteger)
 tdf.Fields.Append fldI092

 Dim fldI093 As DAO.Field
 Set fldI093 = tdf.CreateField("bmi_caf", dbInteger)
 tdf.Fields.Append fldI093

 Dim fldI094 As DAO.Field
 Set fldI094 = tdf.CreateField("dental_prob", dbBoolean)
 tdf.Fields.Append fldI094
 Call SetPropertyDAO(fldI094, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldI095 As DAO.Field
 Set fldI095 = tdf.CreateField("app_down", dbBoolean)
 tdf.Fields.Append fldI095
 Call SetPropertyDAO(fldI095, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldI096 As DAO.Field
 Set fldI096 = tdf.CreateField("ED_indic", dbBoolean)
 tdf.Fields.Append fldI096
 Call SetPropertyDAO(fldI096, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldI097 As DAO.Field
 Set fldI097 = tdf.CreateField("diet_mod_rec", dbBoolean)
 tdf.Fields.Append fldI097
 Call SetPropertyDAO(fldI097, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldI098 As DAO.Field
 Set fldI098 = tdf.CreateField("heath_direc_cur", dbBoolean)
 tdf.Fields.Append fldI098
 Call SetPropertyDAO(fldI098, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldI099 As DAO.Field
 Set fldI099 = tdf.CreateField("health_direc_fut", dbBoolean)
 tdf.Fields.Append fldI099
 Call SetPropertyDAO(fldI099, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldI100 As DAO.Field
 Set fldI100 = tdf.CreateField("appearance", dbText)
 tdf.Fields.Append fldI100
 Call setComboMultiProperties(fldI100, "Well groomed;Disheveled;Bizarre;Inappropriate")

 Dim fldI101 As DAO.Field
 Set fldI101 = tdf.CreateField("attitude", dbText)
 tdf.Fields.Append fldI101
 Call setComboMultiProperties(fldI101, "Cooperative;Guarded;Suspicious;Uncooperative;Hostile;Other")

 Dim fldI102 As DAO.Field
 Set fldI102 = tdf.CreateField("speech", dbText)
 tdf.Fields.Append fldI102
 Call setComboMultiProperties(fldI102, "Normal;Delayed;Soft;Loud;Slurred;Pressured")

 Dim fldI103 As DAO.Field
 Set fldI103 = tdf.CreateField("affect", dbText)
 tdf.Fields.Append fldI103
 Call setComboMultiProperties(fldI103, "Appropriate;Labile;Expansive;Constricted;Blunted;Flat")

 Dim fldI104 As DAO.Field
 Set fldI104 = tdf.CreateField("mood", dbText)
 tdf.Fields.Append fldI104
 Call setComboMultiProperties(fldI104, "Euthymic;Depressed;Agitated;Euphoric;Irritable;Anxious")

 Dim fldI105 As DAO.Field
 Set fldI105 = tdf.CreateField("thought_pro", dbText)
 tdf.Fields.Append fldI105
 Call setComboMultiProperties(fldI105, "Goal-directed;Circumstantial;Tangential;LOA;FOI;Disorganized;Bizarre")

 Dim fldI106 As DAO.Field
 Set fldI106 = tdf.CreateField("thought_cont", dbText)
 tdf.Fields.Append fldI106
 Call setComboMultiProperties(fldI106, "WNL;Paranoia;Phobias;Ideas of reference;Delusions;Obsessions")

 Dim fldI107 As DAO.Field
 Set fldI107 = tdf.CreateField("perception", dbText)
 tdf.Fields.Append fldI107
 Call setComboMultiProperties(fldI107, "WNL;Auditory hallucinations;Visual hallucinations;Illusions")

 Dim fldI108 As DAO.Field
 Set fldI108 = tdf.CreateField("orient", dbText)
 tdf.Fields.Append fldI108
 Call setComboMultiProperties(fldI108, "Person;Place;Date;Other situation;Not oriented")

 Dim fldI109 As DAO.Field
 Set fldI109 = tdf.CreateField("judg", dbText)
 tdf.Fields.Append fldI109
 Call setComboMultiProperties(fldI109, "Intact;Impaired")

 Dim fldI110 As DAO.Field
 Set fldI110 = tdf.CreateField("insight", dbText)
 tdf.Fields.Append fldI110
 Call setComboMultiProperties(fldI110, "Superior;Good;Fair;Poor")

 Dim fldI111 As DAO.Field
 Set fldI111 = tdf.CreateField("fund_know", dbText)
 tdf.Fields.Append fldI111
 Call setComboMultiProperties(fldI111, "Intact;Above average;Average;Below average;Mini mental")

 Dim fldI112 As DAO.Field
 Set fldI112 = tdf.CreateField("lang_mem", dbText)
 tdf.Fields.Append fldI112
 Call setComboProperties(fldI112, "Can name three objects;repeat a phrase;repeats objects correctly;Repeats objects after 5 minutes")

 Dim fldI113 As DAO.Field
 Set fldI113 = tdf.CreateField("lett_for", dbText)
 tdf.Fields.Append fldI113
 Call setComboProperties(fldI113, "correctly;with difficulty;unable")

 Dim fldI113a As DAO.Field
 Set fldI113a = tdf.CreateField("lett_back", dbText)
 tdf.Fields.Append fldI113a
 Call setComboProperties(fldI113a, "correctly;with difficulty;unable")

 Dim fldI114 As DAO.Field
 Set fldI114 = tdf.CreateField("num_for", dbText)
 tdf.Fields.Append fldI114
 Call setComboProperties(fldI114, "correctly;with difficulty;unable")

 Dim fldI115 As DAO.Field
 Set fldI115 = tdf.CreateField("num_back", dbText)
 tdf.Fields.Append fldI115
 Call setComboProperties(fldI115, "correctly;with difficulty;unable")

 Dim fldI116 As DAO.Field
 Set fldI116 = tdf.CreateField("sleep_", dbText)
 tdf.Fields.Append fldI116
 Call setComboProperties(fldI116, "WNL;DFA;MNA;EMA;Increased;Decreased")

 Dim fldI117 As DAO.Field
 Set fldI117 = tdf.CreateField("appetite_", dbText)
 tdf.Fields.Append fldI117
 Call setComboProperties(fldI117, "Normal;Increased;Decreased")

 Dim fldI118 As DAO.Field
 Set fldI118 = tdf.CreateField("energy_", dbText)
 tdf.Fields.Append fldI118
 Call setComboProperties(fldI118, "WNL;Low;High")

 Dim fldI119 As DAO.Field
 Set fldI119 = tdf.CreateField("mh_stage_change", dbText)
 tdf.Fields.Append fldI119

 Dim fldI120 As DAO.Field
 Set fldI120 = tdf.CreateField("sa_stage_change", dbText)
 tdf.Fields.Append fldI120

 Dim fldI121 As DAO.Field
 Set fldI121 = tdf.CreateField("barriers", dbText)
 tdf.Fields.Append fldI121
 Call setComboMultiProperties(fldI121, "None;Language barrier;Psychiatric symptoms;Cognitive deficits;Visual impairment;Hearing impairment;Unable to read;Unable to write;Other")

 Dim fldI122 As DAO.Field
 Set fldI122 = tdf.CreateField("strengths_num", dbInteger)
 tdf.Fields.Append fldI122

 'CAF II fields
 Dim fldJ01 As DAO.Field
 Set fldJ01 = tdf.CreateField("cult_con_life", dbInteger)
 tdf.Fields.Append fldJ01
 '  Call SetPropertyDAO(fldJ01, "DisplayControl", dbInteger, CInt(acCheckBox))
 Call yesNoNoneQuestions(fldJ01)

 Dim fldJ02 As DAO.Field
 Set fldJ02 = tdf.CreateField("cult_con_tx", dbInteger)
 tdf.Fields.Append fldJ02
 '  Call SetPropertyDAO(fldJ02, "DisplayControl", dbInteger, CInt(acCheckBox))
 Call yesNoNoneQuestions(fldJ02)

 Dim fldJ03 As DAO.Field
 Set fldJ03 = tdf.CreateField("relig_child", dbInteger)
 tdf.Fields.Append fldJ03
 '  Call SetPropertyDAO(fldJ03, "DisplayControl", dbInteger, CInt(acCheckBox))
 Call yesNoNoneQuestions(fldJ03)

 Dim fldJ04 As DAO.Field
 Set fldJ04 = tdf.CreateField("spir_cur", dbInteger)
 tdf.Fields.Append fldJ04
 '  Call SetPropertyDAO(fldJ04, "DisplayControl", dbInteger, CInt(acCheckBox))
 Call yesNoNoneQuestions(fldJ04)

 Dim fldJ05 As DAO.Field
 Set fldJ05 = tdf.CreateField("faith_mem_cur", dbInteger)
 tdf.Fields.Append fldJ05
 '  Call SetPropertyDAO(fldJ05, "DisplayControl", dbInteger, CInt(acCheckBox))
 Call yesNoNoneQuestions(fldJ05)

 Dim fldJ05a As DAO.Field
 Set fldJ05a = tdf.CreateField("med_or_pray", dbInteger)
 tdf.Fields.Append fldJ05a
 '  Call SetPropertyDAO(fldJ05, "DisplayControl", dbInteger, CInt(acCheckBox))
 Call yesNoNoneQuestions(fldJ05a)

 Dim fldJ06 As DAO.Field
 Set fldJ06 = tdf.CreateField("edu_cur", dbText)
 tdf.Fields.Append fldJ06
 Call setComboProperties(fldJ06, "None;in school;in a voc/tech/business")

 Dim fldJ07 As DAO.Field
 Set fldJ07 = tdf.CreateField("Spec_edu", dbInteger)
 tdf.Fields.Append fldJ07
 '  Call SetPropertyDAO(fldJ07, "DisplayControl", dbInteger, CInt(acCheckBox))
 Call yesNoNoneQuestions(fldJ07)

 Dim fldJ08 As DAO.Field
 Set fldJ08 = tdf.CreateField("more_edu", dbText)
 tdf.Fields.Append fldJ08
 '  Call SetPropertyDAO(fldJ08, "DisplayControl", dbInteger, CInt(acCheckBox))
 Call setComboProperties(fldJ06, "Yes;No;Not Sure")

 Dim fldJ09 As DAO.Field
 Set fldJ09 = tdf.CreateField("neur_test", dbInteger)
 tdf.Fields.Append fldJ09
 '  Call SetPropertyDAO(fldJ09, "DisplayControl", dbInteger, CInt(acCheckBox))
 Call yesNoNoneQuestions(fldJ09)

 Dim fldJ10 As DAO.Field
 Set fldJ10 = tdf.CreateField("gamb_screen1", dbText)
 tdf.Fields.Append fldJ10
 Call setComboProperties(fldJ10, "Yes;Yes+Within the past year;No")

 Dim fldJ11 As DAO.Field
 Set fldJ11 = tdf.CreateField("gamb_screen2", dbInteger)
 tdf.Fields.Append fldJ11
 '  Call SetPropertyDAO(fldJ11, "DisplayControl", dbInteger, CInt(acCheckBox))
 Call yesNoNoneQuestions(fldJ11)

 Dim fldJ12 As DAO.Field
 Set fldJ12 = tdf.CreateField("gamb_screen3", dbInteger)
 tdf.Fields.Append fldJ12
 '  Call SetPropertyDAO(fldJ12, "DisplayControl", dbInteger, CInt(acCheckBox))
 Call yesNoNoneQuestions(fldJ12)

 Dim fldJ13 As DAO.Field
 Set fldJ13 = tdf.CreateField("gamb_screen4", dbInteger)
 tdf.Fields.Append fldJ13
 '  Call SetPropertyDAO(fldJ13, "DisplayControl", dbInteger, CInt(acCheckBox))
 Call yesNoNoneQuestions(fldJ13)

 Dim fldJ14 As DAO.Field
 Set fldJ14 = tdf.CreateField("gamb_screen5", dbInteger)
 tdf.Fields.Append fldJ14
 '  Call SetPropertyDAO(fldJ14, "DisplayControl", dbInteger, CInt(acCheckBox))
 Call yesNoNoneQuestions(fldJ14)

 Dim fldJ15 As DAO.Field
 Set fldJ15 = tdf.CreateField("gamb_screen6", dbInteger)
 tdf.Fields.Append fldJ15
 '  Call SetPropertyDAO(fldJ15, "DisplayControl", dbInteger, CInt(acCheckBox))
 Call yesNoNoneQuestions(fldJ15)

 Dim fldJ16 As DAO.Field
 Set fldJ16 = tdf.CreateField("Leg_hx_prob", dbBoolean)
 tdf.Fields.Append fldJ16
 Call SetPropertyDAO(fldJ16, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldJ17 As DAO.Field
 Set fldJ17 = tdf.CreateField("leg_hx_childwef", dbBoolean)
 tdf.Fields.Append fldJ17
 Call SetPropertyDAO(fldJ17, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldJ18 As DAO.Field
 Set fldJ18 = tdf.CreateField("leg_hx_diver", dbBoolean)
 tdf.Fields.Append fldJ18
 Call SetPropertyDAO(fldJ18, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldJ19 As DAO.Field
 Set fldJ19 = tdf.CreateField("leg_hx_parole", dbBoolean)
 tdf.Fields.Append fldJ19
 Call SetPropertyDAO(fldJ19, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldJ20 As DAO.Field
 Set fldJ20 = tdf.CreateField("leg_hx_civil", dbBoolean)
 tdf.Fields.Append fldJ20
 Call SetPropertyDAO(fldJ20, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldJ21 As DAO.Field
 Set fldJ21 = tdf.CreateField("leg_hx_asist", dbBoolean)
 tdf.Fields.Append fldJ21
 Call SetPropertyDAO(fldJ21, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldJ22 As DAO.Field
 Set fldJ22 = tdf.CreateField("leg_hx_psrb", dbBoolean)
 tdf.Fields.Append fldJ22
 Call SetPropertyDAO(fldJ22, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldJ23 As DAO.Field
 Set fldJ23 = tdf.CreateField("leg_hx_famcourt", dbBoolean)
 tdf.Fields.Append fldJ23
 Call SetPropertyDAO(fldJ23, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldJ24 As DAO.Field
 Set fldJ24 = tdf.CreateField("leg_hx_crest", dbBoolean)
 tdf.Fields.Append fldJ24
 Call SetPropertyDAO(fldJ24, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldJ25 As DAO.Field
 Set fldJ25 = tdf.CreateField("leg_hx_54", dbBoolean)
 tdf.Fields.Append fldJ25
 Call SetPropertyDAO(fldJ25, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldJ26 As DAO.Field
 Set fldJ26 = tdf.CreateField("leg_hx_crimcourt", dbBoolean)
 tdf.Fields.Append fldJ26
 Call SetPropertyDAO(fldJ26, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldJ27 As DAO.Field
 Set fldJ27 = tdf.CreateField("leg_hx_corp", dbBoolean)
 tdf.Fields.Append fldJ27
 Call SetPropertyDAO(fldJ27, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldJ28 As DAO.Field
 Set fldJ28 = tdf.CreateField("leg_hx_protorder", dbBoolean)
 tdf.Fields.Append fldJ28
 Call SetPropertyDAO(fldJ28, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldJ29 As DAO.Field
 Set fldJ29 = tdf.CreateField("Arrested2", dbBoolean)
 tdf.Fields.Append fldJ29
 Call SetPropertyDAO(fldJ29, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldJ30 As DAO.Field
 Set fldJ30 = tdf.CreateField("cur_mand_mh", dbText)
 tdf.Fields.Append fldJ30
 Call setComboProperties(fldJ30, "Yes;No;Unknown")

 Dim fldJ31 As DAO.Field
 Set fldJ31 = tdf.CreateField("cur_mand_sub", dbText)
 tdf.Fields.Append fldJ31
 Call setComboProperties(fldJ31, "Yes;No;Unknown")

 Dim fldJ32 As DAO.Field
 Set fldJ32 = tdf.CreateField("traum_critA", dbText)
 tdf.Fields.Append fldJ32
 Call setComboMultiProperties(fldJ32, "Yes;No;Now/Recent")

 Dim fldJ33 As DAO.Field
 Set fldJ33 = tdf.CreateField("traum_event1", dbText)
 tdf.Fields.Append fldJ33
 Call setComboMultiProperties(fldJ33, "Yes;No;Now/Recent")

 Dim fldJ34 As DAO.Field
 Set fldJ34 = tdf.CreateField("traum_event2", dbText)
 tdf.Fields.Append fldJ34
 Call setComboMultiProperties(fldJ34, "Yes;No;Now/Recent")

 Dim fldJ35 As DAO.Field
 Set fldJ35 = tdf.CreateField("traum_event3", dbText)
 tdf.Fields.Append fldJ35
 Call setComboMultiProperties(fldJ35, "Yes;No;Now/Recent")

 Dim fldJ36 As DAO.Field
 Set fldJ36 = tdf.CreateField("traum_event4", dbText)
 tdf.Fields.Append fldJ36
 Call setComboMultiProperties(fldJ36, "Yes;No;Now/Recent")

 Dim fldJ37 As DAO.Field
 Set fldJ37 = tdf.CreateField("traum_event5", dbText)
 tdf.Fields.Append fldJ37
 Call setComboMultiProperties(fldJ37, "Yes;No;Now/Recent")

 Dim fldJ38 As DAO.Field
 Set fldJ38 = tdf.CreateField("traum_event6", dbText)
 tdf.Fields.Append fldJ38
 Call setComboMultiProperties(fldJ38, "Yes;No;Now/Recent")

 Dim fldJ39 As DAO.Field
 Set fldJ39 = tdf.CreateField("traum_event7", dbText)
 tdf.Fields.Append fldJ39
 Call setComboMultiProperties(fldJ39, "Yes;No;Now/Recent")

 Dim fldJ40 As DAO.Field
 Set fldJ40 = tdf.CreateField("traum_event8", dbText)
 tdf.Fields.Append fldJ40
 Call setComboMultiProperties(fldJ40, "Yes;No;Now/Recent")

 Dim fldJ41 As DAO.Field
 Set fldJ41 = tdf.CreateField("traum_event9", dbText)
 tdf.Fields.Append fldJ41
 Call setComboMultiProperties(fldJ41, "Yes;No;Now/Recent")

 Dim fldJ42 As DAO.Field
 Set fldJ42 = tdf.CreateField("traum_event10", dbText)
 tdf.Fields.Append fldJ42
 Call setComboMultiProperties(fldJ42, "Yes;No;Now/Recent")

 Dim fldJ43 As DAO.Field
 Set fldJ43 = tdf.CreateField("traum_event_tx_conc", dbBoolean)
 tdf.Fields.Append fldJ43
 Call SetPropertyDAO(fldJ43, "DisplayControl", dbInteger, CInt(acCheckBox))

 Dim fldJ44 As DAO.Field
 Set fldJ44 = tdf.CreateField("unsafe", dbBoolean)
 tdf.Fields.Append fldJ44
 Call SetPropertyDAO(fldJ44, "DisplayControl", dbInteger, CInt(acCheckBox))
  
 'refresh the tables and database
 db.TableDefs.Refresh
 Application.RefreshDatabaseWindow
 
 Debug.Print "Done"

End Sub

Function setComboProperties(obj As Object, strList As String)
    
 With obj
  .Properties.Append .CreateProperty("DisplayControl", dbInteger, AcControlType.acComboBox)
  .Properties.Append .CreateProperty("RowSourceType", dbText, "Value List")
  .Properties.Append .CreateProperty("RowSource", dbText, strList)
  .Properties.Append .CreateProperty("LimitToList", dbBoolean, True)
 End With

End Function

Function setComboMultiProperties(obj As Object, strList As String)
    
 With obj
  .Properties.Append .CreateProperty("DisplayControl", dbInteger, AcControlType.acComboBox)
  .Properties.Append .CreateProperty("RowSourceType", dbText, "Value List")
  .Properties.Append .CreateProperty("RowSource", dbText, strList)
  .Properties.Append .CreateProperty("LimitToList", dbBoolean, True)
  .Properties.Append .CreateProperty("AllowMultipleValues", dbBoolean, True)
 End With

End Function

Function yesNoNoneQuestions(obj As Object)
 ' Purpose:   Set the properties of ...
 ' Arguments: obj = the object whose property should be set.
    
 Call SetPropertyDAO(obj, "ValidationRule", dbText, "IN (0,1,99)")
 Call SetPropertyDAO(obj, "ValidationText", dbText, "Accepts 0(=no), 1(=yes), or 99(=N/A)")
 Call SetPropertyDAO(obj, "DefaultValue", dbInteger, 0)

End Function

Function intOneToSevenRange(obj As Object)
 ' Purpose:   Set the properties of ...
 ' Arguments: obj = the object whose property should be set.
    
 Call SetPropertyDAO(obj, "ValidationRule", dbText, "IN (1,2,3,4,5,6,7)")
 Call SetPropertyDAO(obj, "ValidationText", dbText, "Accepts integers 1 through 7")

End Function

Function SetPropertyDAO(obj As Object, strPropertyName As String, intType As Integer, _
    varValue As Variant, Optional strErrMsg As String) As Boolean
 On Error GoTo ErrHandler
    'Purpose:   Set a property for an object, creating if necessary.
    'Arguments: obj = the object whose property should be set.
    '           strPropertyName = the name of the property to set.
    '           intType = the type of property (needed for creating)
    '           varValue = the value to set this property to.
    '           strErrMsg = string to append any error message to.
    
 If HasProperty(obj, strPropertyName) Then
  obj.Properties(strPropertyName) = varValue
 Else
  obj.Properties.Append obj.CreateProperty(strPropertyName, intType, varValue)
 End If
 SetPropertyDAO = True
 
ExitHandler:
 Exit Function

ErrHandler:
 strErrMsg = strErrMsg & obj.Name & "." & strPropertyName & " not set to " & varValue & _
  ". Error " & Err.Number & " - " & Err.Description & vbCrLf
 Resume ExitHandler
End Function

Public Function HasProperty(obj As Object, strPropName As String) As Boolean
 'Purpose:   Return true if the object has the property.
 Dim varDummy As Variant
 
 On Error Resume Next
 varDummy = obj.Properties(strPropName)
 HasProperty = (Err.Number = 0)
End Function



