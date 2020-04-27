using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using TealCompetancy.App_Start;
using TealCompetancy.Models;
using WhiteGod;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Data;

/*
Controller of the Competency extractor Module
Developer : C Vellaichamy
Date: Feb 2020
*/



namespace TealCompetancy.Controllers
{
    [AuthorizationPrivilegeFilter]
    public class HomeController : Controller
    {
        // GET: Home
        General CtrlGen = new General();
        BO_General BOGen = new BO_General();
        DA_DBCon DBCon = new DA_DBCon();
        public ActionResult Index()
        {
            if (Convert.ToString(Session["SInfoMsg"]) != "")
            {
                CtrlGen.Message_Info = CommonCls.GetMsg(Convert.ToString(Session["SInfoMsg"]).Trim(), Convert.ToString(Session["SInfoMsgType"]).Trim());
                Session["SInfoMsg"] = "";
                Session["SInfoMsgType"] = "";
            }
            return View(CtrlGen);
        }
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Index(General ObjValue)
        {
            try
            {
                Session["SInfoMsg"] = "";
                Session["SInfoMsgType"] = "";
                if (CheckRequiredField(ObjValue) == false) // Check Server Side Required Field Validation
                {
                    Session["SInfoMsg"] = "<ul style='padding-left: 10px;padding-right: 10px;'>" + Convert.ToString(Session["SInfoMsg"]).Trim() + "</ul>";
                    ObjValue.Message_Info = CommonCls.GetMsg(Convert.ToString(Session["SInfoMsg"]).Trim(), Convert.ToString(Session["SInfoMsgType"]).Trim());
                    Session["SInfoMsg"] = "";
                    Session["SInfoMsgType"] = "";
                }
                else
                {
                    if (ObjValue.competancy_file != null)
                    {
                        // Get file and save path Start
                        string StrFilePath = Server.MapPath("~/Content/CDSUploadFile").Replace("/", @"\").Replace("\\", @"\");
                        White.CreateDirectory(StrFilePath);
                        string[] StrReplace = { " ", "," };
                        string fileName = Path.GetFileNameWithoutExtension(ObjValue.competancy_file.FileName).Trim();
                        string StrExtension = Path.GetExtension(ObjValue.competancy_file.FileName).Trim();
                        fileName = White.ReplaceWords(fileName, StrReplace, "_") + DateTime.Now.ToString("ddMMyyyyHHmm") + StrExtension;
                        string StrFilePathAndName = StrFilePath + "/" + fileName;
                        ObjValue.competancy_file.SaveAs(StrFilePathAndName);
                        // Get file and save path End
                        // Excel validation start
                        Boolean BlnResult = ValidationXLSX(ObjValue, StrFilePathAndName);
                        // Excel validation End
                        if (BlnResult == true)
                        {
                            if(DataValidation() == true)
                            {
                                Session["SInfoMsg"] = "<ul style='padding-left: 10px;padding-right: 10px;'>" + Convert.ToString(Session["SInfoMsg"]).Trim() + "</ul>";
                                ObjValue.Message_Info = CommonCls.GetMsg(Convert.ToString(Session["SInfoMsg"]).Trim(), Convert.ToString(Session["SInfoMsgType"]).Trim());
                                Session["SInfoMsg"] = "";
                                Session["SInfoMsgType"] = "";
                            }
                            else
                            {
                                // Data Validation failed
                                Session["SInfoMsg"] = "<ul style='padding-left: 10px;padding-right: 10px;'>" + Convert.ToString(Session["SInfoMsg"]).Trim() + "</ul>";
                                ObjValue.Message_Info = CommonCls.GetMsg(Convert.ToString(Session["SInfoMsg"]).Trim(), Convert.ToString(Session["SInfoMsgType"]).Trim());
                                Session["SInfoMsg"] = "";
                                Session["SInfoMsgType"] = "";
                                CommonCls.FileDelete(StrFilePathAndName);
                            }
                        }
                        else
                        {
                            // Excel Validation failed
                            Session["SInfoMsg"] = "<ul style='padding-left: 10px;padding-right: 10px;'>" + Convert.ToString(Session["SInfoMsg"]).Trim() + "</ul>";
                            ObjValue.Message_Info = CommonCls.GetMsg(Convert.ToString(Session["SInfoMsg"]).Trim(), Convert.ToString(Session["SInfoMsgType"]).Trim());
                            Session["SInfoMsg"] = "";
                            Session["SInfoMsgType"] = "";
                            CommonCls.FileDelete(StrFilePathAndName);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                logger.write("HomeController.cs", "ERROR ON Index(ObjValue) - " + ex.Message);
                ObjValue.Message_Info = CommonCls.GetMsg(ex.Message, "ERROR");
            }
            return View(ObjValue);
        }
        #region "Download"
        public FileResult Download()
        {
            string strAbsolutePath = string.Format("{0}/{1}", HttpRuntime.AppDomainAppPath + "Content", "doc");
            strAbsolutePath = strAbsolutePath.Replace("/", @"\").Replace("\\", @"\");
            strAbsolutePath += "\\Sample Sheet.xlsx";
            byte[] fileBytes = System.IO.File.ReadAllBytes(strAbsolutePath);
            string fileName = "Sample Sheet.xlsx";
            return File(fileBytes, System.Net.Mime.MediaTypeNames.Application.Octet, fileName);
        }
        #endregion
        #region "Data Validation"
        private Boolean DataValidation()
        {
            Boolean BlnResult = true;
            string StrError = "";
            try
            {
                int IntError = 0;
                DataTable dtSheetList = new DataTable();
                DataTable dtEmpList = new DataTable();
                DataTable dtCompList = new DataTable();
                DataTable dtReqList = new DataTable();
                // DataTable Sesstion to DataTable Convert start
                dtSheetList = (DataTable)Session["Sheet_List"];
                dtEmpList = (DataTable)Session["Emp_List"];
                dtCompList = (DataTable)Session["Comp_List"];
                dtReqList = (DataTable)Session["Req_List"];
                // DataTable Sesstion to DataTable Convert end
                if (dtEmpList.Rows.Count > 0)
                {
                    DataTable dt = new DataTable();
                    BOGen.emp_code = Convert.ToString(dtEmpList.Rows[0]["emp_code"]).Trim();
                    dt = BOGen.BO_get_emp_details_select_with_emp_code();
                    if (dt.Rows.Count > 0) // Check exists employee
                    {
                        StrError += CommonCls.GetErrorMsg("["+ Convert.ToString(dtEmpList.Rows[0]["emp_code"]).Trim() + "] - this employee code already exists.");
                        BlnResult = false;
                        IntError = 1;
                    }
                    DataTable distinctTable = dtCompList.DefaultView.ToTable(true, "comp_desc");
                    if (distinctTable.Rows.Count != dtCompList.Rows.Count) // Find Duplicate Records
                    {
                        StrError += CommonCls.GetErrorMsg("Duplicate records found in competency details. Please remove duplicate records.");
                        BlnResult = false;
                        IntError = 1;
                    }
                    if (IntError == 0)
                    {
                        string Strdept_id = "0", Strrole_id = "0";
                        string StrEmp_Code = Convert.ToString(dtEmpList.Rows[0]["emp_code"]).Trim();
                        string Strrole_name = Convert.ToString(dtEmpList.Rows[0]["emp_designation"]).Trim();
                        // Get Department id or Insert Start
                        BOGen.dept_name = Convert.ToString(dtEmpList.Rows[0]["emp_department"]).Trim();
                        dt = BOGen.BO_get_tbl_department_select_with_department_name();
                        if(dt.Rows.Count > 0)
                        {
                            Strdept_id = Convert.ToString(dt.Rows[0]["Dept_id"]).Trim();
                        }
                        else
                        {
                            BOGen.dept_name = Convert.ToString(dtEmpList.Rows[0]["emp_department"]).Trim();
                            Strdept_id = BOGen.BO_tbl_department_insert();
                        }
                        // Get Department id or Insert End
                        // Get Designation id or Insert Start
                        BOGen.role_name = Convert.ToString(dtEmpList.Rows[0]["emp_designation"]).Trim();
                        dt = BOGen.BO_get_tb_role_select_with_role_name();
                        if (dt.Rows.Count > 0)
                        {
                            Strrole_id = Convert.ToString(dt.Rows[0]["role_id"]).Trim();
                        }
                        else
                        {
                            BOGen.role_name = Convert.ToString(dtEmpList.Rows[0]["emp_designation"]).Trim();
                            Strrole_id = BOGen.BO_tb_role_insert();
                        }
                        // Get Designation id or Insert End
                        // Employee Insert Start
                        BOGen.emp_code = StrEmp_Code;
                        BOGen.emp_name = Convert.ToString(dtEmpList.Rows[0]["emp_name"]).Trim();
                        BOGen.emp_total_exp = Convert.ToString(dtEmpList.Rows[0]["emp_exp_yrs_total"]).Trim();
                        BOGen.emp_titan_exp = Convert.ToString(dtEmpList.Rows[0]["emp_exp_yrs_titan"]).Trim();
                        BOGen.emp_outside_exp = Convert.ToString(dtEmpList.Rows[0]["emp_exp_yrs_outside"]).Trim();
                        BOGen.emp_level = Convert.ToString(dtEmpList.Rows[0]["emp_level"]).Trim();
                        BOGen.dept_id = Strdept_id;
                        BOGen.role_id = Strrole_id;
                        BOGen.emp_team_lead = Convert.ToString(dtEmpList.Rows[0]["emp_supervisor"]).Trim();
                        BOGen.BO_emp_details_insert_update();
                        // Employee Insert End
                        // Competency Insert Start
                        if (dtCompList.Rows.Count > 0)
                        {
                            for (int i = 0; i < dtCompList.Rows.Count; i++)
                            {
                                // role_comp_map Insert
                                BOGen.role_name = Strrole_name;
                                BOGen.comp_name = Convert.ToString(dtCompList.Rows[i]["comp_desc"]).Trim();
                                BOGen.BO_role_comp_map_insert();

                                // settings_comp Insert
                                BOGen.comp_type = Convert.ToString(dtCompList.Rows[i]["comp_type"]).Trim();
                                BOGen.comp_name = Convert.ToString(dtCompList.Rows[i]["comp_desc"]).Trim();
                                BOGen.required_level = Convert.ToString(dtCompList.Rows[i]["comp_rcl"]).Trim();
                                BOGen.BO_settings_comp_insert();

                                // comp_list Insert
                                BOGen.role_id = Strrole_id;
                                BOGen.emp_code = StrEmp_Code;
                                BOGen.comp_type = Convert.ToString(dtCompList.Rows[i]["comp_type"]).Trim();
                                BOGen.comp_name = Convert.ToString(dtCompList.Rows[i]["comp_desc"]).Trim();
                                BOGen.actual_level = Convert.ToString(dtCompList.Rows[i]["comp_acl"]).Trim();
                                BOGen.required_level = Convert.ToString(dtCompList.Rows[i]["comp_rcl"]).Trim();
                                BOGen.BO_comp_list_insert();

                                DataView dv = new DataView(dtReqList);
                                dv.RowFilter = "req_comp_desc='" + Convert.ToString(dtCompList.Rows[i]["comp_desc"]).Trim() + "'";
                                if(dv.Count > 0)
                                {
                                    for (int j = 0; j < dv.Count; j++)
                                    {
                                        // settings_skill Insert
                                        BOGen.comp_name = Convert.ToString(dtCompList.Rows[i]["comp_desc"]).Trim();
                                        BOGen.applicable_levels = Convert.ToString(dv[j]["applicable_levels"]).Trim();
                                        BOGen.skill_req = Convert.ToString(dv[j]["requirements"]).Trim();
                                        BOGen.knowledge_scale = Convert.ToString(dv[j]["knowledge_scale"]).Trim();
                                        BOGen.BO_settings_skill_insert();

                                        // tb_skill Insert
                                        int int_scale = 0;
                                        if(Convert.ToString(dv[j]["knowledge_scale"]).ToUpper().Trim() != "YET TO START")
                                        {
                                            string strvalue = White.Right(Convert.ToString(dv[j]["knowledge_scale"]).Trim(), 1);
                                            int_scale = strvalue == "" ? 0 : Convert.ToInt32(strvalue);
                                        }
                                        BOGen.emp_code = StrEmp_Code;
                                        BOGen.comp_name = Convert.ToString(dtCompList.Rows[i]["comp_desc"]).Trim();
                                        BOGen.applicable_levels = Convert.ToString(dv[j]["applicable_levels"]).Trim();
                                        BOGen.skill_req = Convert.ToString(dv[j]["requirements"]).Trim();
                                        BOGen.knowledge_scale = Convert.ToString(dv[j]["knowledge_scale"]).Trim();
                                        BOGen.int_scale = Convert.ToString(int_scale).Trim();
                                        BOGen.BO_tb_skill_insert();
                                    }
                                }
                            }
                        }
                        // Competency Insert End
                    }
                }
                else
                {
                    StrError += CommonCls.GetErrorMsg("Competency employee details not found in summary sheet.");
                    BlnResult = false;
                    IntError = 1;
                }
                if (StrError != "")
                {
                    Session["SInfoMsg"] = StrError;
                    Session["SInfoMsgType"] = "ERROR";
                }
                else
                {
                    Session["SInfoMsg"] = "File uploaded successfully...";
                    Session["SInfoMsgType"] = "INFO";
                }
            }
            catch (Exception ex)
            {
                BlnResult = false;
                logger.write("HomeController.cs", "ERROR ON DataValidation() - " + ex.Message);
                Session["SInfoMsg"] = ex.Message;
                Session["SInfoMsgType"] = "ERROR";

            }
            return BlnResult;
        }

        #endregion
        #region "XLSX Validation"
        private bool ValidationXLSX(General ObjValue, string StrFilePathAndName)
        {
            Boolean BlnResult = true;
            string StrError = "";
            try
            {
                if (StrFilePathAndName.Trim() == "")
                {
                    StrError += CommonCls.GetErrorMsg("Competency data sheet upload failed...");
                    BlnResult = false;
                }
                else
                {
                    // Initial DataTable Load Start
                    SetInitialRow();
                    // Initial DataTable Load End
                    int IntError = 0;
                    DataTable dtSheetList = new DataTable();
                    DataTable dtEmpList = new DataTable();
                    DataTable dtCompList = new DataTable();
                    DataTable dtReqList = new DataTable();
                    // Session DataTable to DataTable Convert Start
                    dtSheetList = (DataTable)Session["Sheet_List"];
                    dtEmpList = (DataTable)Session["Emp_List"];
                    dtCompList = (DataTable)Session["Comp_List"];
                    dtReqList = (DataTable)Session["Req_List"];
                    // Session DataTable to DataTable Convert End
                    using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(StrFilePathAndName, false))
                    {
                        WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                        Sheet SummarySheet = workbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name.ToString().ToUpper() == "SUMMARY").FirstOrDefault();
                        if (SummarySheet == null) // Sheet Name validation
                        {
                            StrError += CommonCls.GetErrorMsg("Summary sheet not found in the uploaded file. Please upload the competency data sheet in the correct format.");
                            BlnResult = false;
                            IntError = 1;
                        }
                        else
                        {
                            // Get Sheet Names Start
                            var workbook = spreadsheetDocument.WorkbookPart.Workbook;
                            var AllSheets = workbook.Sheets.Cast<Sheet>().ToList();
                            foreach (Sheet CtrlSheet in AllSheets)
                            {
                                DataRow dr = dtSheetList.NewRow();
                                dr["Relationship_Id"] = CtrlSheet.Id.Value;
                                dr["Sheet_Id"] = CtrlSheet.SheetId.Value;
                                dr["Sheet_Name"] = CtrlSheet.Name.Value;
                                dtSheetList.Rows.Add(dr);
                                dtSheetList.AcceptChanges();
                            }
                            Session["Sheet_List"] = dtSheetList;
                            // Get Sheet Names End
                            // Get Summary Sheet Values Start
                            string StrColEmpName = ClsExcel.GetCellValue(workbookPart, SummarySheet, "A1");
                            string StrColEmpCode = ClsExcel.GetCellValue(workbookPart, SummarySheet, "A2");
                            string StrColEmpExpYrsTot = ClsExcel.GetCellValue(workbookPart, SummarySheet, "A3");
                            string StrColEmpExpYrsTit = ClsExcel.GetCellValue(workbookPart, SummarySheet, "A4");
                            string StrColEmpExpYrsOutSide = ClsExcel.GetCellValue(workbookPart, SummarySheet, "A5");

                            string StrColEmpLevel = ClsExcel.GetCellValue(workbookPart, SummarySheet, "C1");
                            string StrColEmpDept = ClsExcel.GetCellValue(workbookPart, SummarySheet, "C2");
                            string StrColEmpDes = ClsExcel.GetCellValue(workbookPart, SummarySheet, "C3");
                            string StrColEmpSup = ClsExcel.GetCellValue(workbookPart, SummarySheet, "C4");
                            string StrColEmpReviewer = ClsExcel.GetCellValue(workbookPart, SummarySheet, "C5");

                            string StrColCompType = ClsExcel.GetCellValue(workbookPart, SummarySheet, "A8");
                            string StrColCompDesc = ClsExcel.GetCellValue(workbookPart, SummarySheet, "B8");
                            string StrColACL = ClsExcel.GetCellValue(workbookPart, SummarySheet, "C8");
                            string StrColRCL = ClsExcel.GetCellValue(workbookPart, SummarySheet, "D8");

                            if (StrColEmpName.ToUpper().Trim() == "NAME" && StrColEmpCode.ToUpper().Trim() == "EMP CODE" && StrColEmpExpYrsTot.ToUpper().Trim() == "EXP. IN YRS(TOTAL)" && StrColEmpExpYrsTit.ToUpper().Trim() == "EXP. IN YRS(TITAN)" && StrColEmpExpYrsOutSide.ToUpper().Trim() == "EXP. IN YRS(OUTSIDE)" && StrColEmpLevel.ToUpper().Trim() == "LEVEL" && StrColEmpDept.ToUpper().Trim() == "DEPARTMENT" && StrColEmpDes.ToUpper().Trim() == "DESIGNATION" && StrColEmpSup.ToUpper().Trim() == "SUPERVISOR" && StrColEmpReviewer.ToUpper().Trim() == "REVIEWER" && StrColCompType.ToUpper().Trim() == "COMPETENCY TYPE" && StrColCompDesc.ToUpper().Trim() == "COMPETENCY DESCRIPTION" && StrColACL.ToUpper().Trim() == "ACTUAL COMPETENCY LEVEL" && StrColRCL.ToUpper().Trim() == "REQUIRED COMPETENCY LEVEL")
                            {
                                string StrEmpName = ClsExcel.GetCellValue(workbookPart, SummarySheet, "B1");
                                string StrEmpCode = ClsExcel.GetCellValue(workbookPart, SummarySheet, "B2");
                                string StrEmpExpYrsTot = ClsExcel.GetCellValue(workbookPart, SummarySheet, "B3");
                                string StrEmpExpYrsTit = ClsExcel.GetCellValue(workbookPart, SummarySheet, "B4");
                                string StrEmpExpYrsOutSide = ClsExcel.GetCellValue(workbookPart, SummarySheet, "B5");

                                string StrEmpLevel = ClsExcel.GetCellValue(workbookPart, SummarySheet, "D1");
                                string StrEmpDept = ClsExcel.GetCellValue(workbookPart, SummarySheet, "D2");
                                string StrEmpDes = ClsExcel.GetCellValue(workbookPart, SummarySheet, "D3");
                                string StrEmpSup = ClsExcel.GetCellValue(workbookPart, SummarySheet, "D4");
                                string StrEmpReviewer = ClsExcel.GetCellValue(workbookPart, SummarySheet, "D5");
                                if(EmpCheckRequiredField(StrEmpName,StrEmpCode,StrEmpExpYrsTot,StrEmpExpYrsTit,StrEmpExpYrsOutSide,StrEmpLevel,StrEmpDept,StrEmpDes,StrEmpSup,StrEmpReviewer) == false)
                                {
                                    StrError += Convert.ToString(Session["SInfoMsg"]).Trim();
                                    Session["SInfoMsg"] = "";
                                    Session["SInfoMsgType"] = "";
                                    BlnResult = false;
                                    IntError = 1;
                                }
                                if(CommonCls.IsDecimalNumber(StrEmpExpYrsTot) == false)
                                {
                                    StrError += CommonCls.GetErrorMsg("Exp. In Yrs(Total) is invalid format value.");
                                    Session["SInfoMsg"] = "";
                                    Session["SInfoMsgType"] = "";                                            
                                    BlnResult = false;
                                    IntError = 1;
                                }
                                if (CommonCls.IsDecimalNumber(StrEmpExpYrsTit) == false)
                                {
                                    StrError += CommonCls.GetErrorMsg("Exp. In Yrs(Titan) is invalid format value.");
                                    Session["SInfoMsg"] = "";
                                    Session["SInfoMsgType"] = "";
                                    BlnResult = false;
                                    IntError = 1;
                                }
                                if (CommonCls.IsDecimalNumber(StrEmpExpYrsOutSide) == false)
                                {
                                    StrError += CommonCls.GetErrorMsg("Exp. In Yrs(Outside) is invalid format value.");
                                    Session["SInfoMsg"] = "";
                                    Session["SInfoMsgType"] = "";
                                    BlnResult = false;
                                    IntError = 1;
                                }
                                DataRow dr = dtEmpList.NewRow();
                                dr["emp_name"] = StrEmpName;
                                dr["emp_code"] = StrEmpCode;
                                dr["emp_exp_yrs_total"] = StrEmpExpYrsTot;
                                dr["emp_exp_yrs_titan"] = StrEmpExpYrsTit;
                                dr["emp_exp_yrs_outside"] = StrEmpExpYrsOutSide;
                                dr["emp_level"] = StrEmpLevel;
                                dr["emp_department"] = StrEmpDept;
                                dr["emp_designation"] = StrEmpDes;
                                dr["emp_supervisor"] = StrEmpSup;
                                dr["emp_reviewer"] = StrEmpReviewer;
                                dtEmpList.Rows.Add(dr);
                                dtEmpList.AcceptChanges();
                                Session["Emp_List"] = dtEmpList;
                                int IntColCount = 8, IntEmptyCount = 0;
                                while (IntEmptyCount < 5)
                                {
                                    IntColCount += 1;
                                    string StrCompType = ClsExcel.GetCellValue(workbookPart, SummarySheet, "A" + IntColCount);
                                    string StrCompDesc = ClsExcel.GetCellValue(workbookPart, SummarySheet, "B" + IntColCount);
                                    string StrACL = ClsExcel.GetCellValue(workbookPart, SummarySheet, "C" + IntColCount);
                                    string StrRCL = ClsExcel.GetCellValue(workbookPart, SummarySheet, "D" + IntColCount);
                                    if(StrCompType.Trim() != "" || StrCompDesc.Trim() != "" || StrACL.Trim() != "" || StrRCL.Trim() != "")
                                    {
                                        IntEmptyCount = 0;
                                        DataRow drComp = dtCompList.NewRow();
                                        drComp["comp_type"] = StrCompType;
                                        drComp["comp_desc"] = StrCompDesc;
                                        drComp["comp_acl"] = StrACL;
                                        drComp["comp_rcl"] = StrRCL;
                                        dtCompList.Rows.Add(drComp);
                                        dtCompList.AcceptChanges();
                                        if(CompCheckRequiredField(StrCompType,StrCompDesc,StrACL,StrRCL) == false)
                                        {
                                            StrError += Convert.ToString(Session["SInfoMsg"]).Trim();
                                            Session["SInfoMsg"] = "";
                                            Session["SInfoMsgType"] = "";
                                            BlnResult = false;
                                            IntError = 1;
                                        }
                                    }
                                    else
                                    {
                                        IntEmptyCount += 1;
                                    }
                                }
                                Session["Comp_List"] = dtCompList;
                                if(dtCompList.Rows.Count <= 0)
                                {
                                    StrError += CommonCls.GetErrorMsg("Competency details not found in summary sheet.");
                                    BlnResult = false;
                                    IntError = 1;
                                }
                            }
                            else
                            {
                                StrError += CommonCls.GetErrorMsg("Summary sheet format is invalid. Please upload the competency data sheet in the correct format.");
                                BlnResult = false;
                                IntError = 1;
                            }
                            // Get Summary Sheet Values End
                            // Other sheet find Start
                            if (IntError == 0)
                            {
                                if(dtCompList.Rows.Count > 0)
                                {
                                    for (int i = 0; i < dtCompList.Rows.Count; i++)
                                    {
                                        Sheet FindSheet = workbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name.ToString().ToUpper() == Convert.ToString(dtCompList.Rows[i]["comp_desc"]).ToUpper().Trim()).FirstOrDefault();
                                        if(FindSheet == null)
                                        {
                                            StrError +=  CommonCls.GetErrorMsg("&quot;" + Convert.ToString(dtCompList.Rows[i]["comp_desc"]).Trim() + "&quot; sheet not found. Please add the &quot;" + Convert.ToString(dtCompList.Rows[i]["comp_desc"]).Trim() + "&quot; sheet.");
                                            BlnResult = false;
                                            IntError = 1;
                                        }
                                    }
                                }
                            }
                            // Other sheet find End
                            // Other sheet validation Start
                            if (IntError == 0)
                            {
                                if (dtCompList.Rows.Count > 0)
                                {
                                    for (int i = 0; i < dtCompList.Rows.Count; i++)
                                    {
                                        Sheet CtrlSheet = workbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name.ToString().ToUpper() == Convert.ToString(dtCompList.Rows[i]["comp_desc"]).ToUpper().Trim()).FirstOrDefault();
                                        if (CtrlSheet != null)
                                        {
                                            string StrColSLNo = ClsExcel.GetCellValue(workbookPart, CtrlSheet, "A1");
                                            string StrColAL = ClsExcel.GetCellValue(workbookPart, CtrlSheet, "B1");
                                            string StrColReq = ClsExcel.GetCellValue(workbookPart, CtrlSheet, "C1");
                                            string StrColKS = ClsExcel.GetCellValue(workbookPart, CtrlSheet, "E1");
                                            if (StrColSLNo.ToUpper().Trim() != "SL NO" || StrColAL.ToUpper().Trim() != "APPLICABLE LEVELS" || StrColReq.ToUpper().Trim() != "REQUIREMENTS" || StrColKS.ToUpper().Trim() != "KNOWLEDGE SCALE")
                                            {
                                                StrError += CommonCls.GetErrorMsg("&quot;" + Convert.ToString(dtCompList.Rows[i]["comp_desc"]).Trim() + "&quot; sheet format is invalid. Please upload the &quot;" + Convert.ToString(dtCompList.Rows[i]["comp_desc"]).Trim() + "&quot; sheet in the correct format.");
                                                BlnResult = false;
                                                IntError = 1;
                                            }
                                        }
                                    }
                                }
                            }
                            // Other sheet validation End
                            // Other sheet required validation Start
                            if (IntError == 0)
                            {
                                if (dtCompList.Rows.Count > 0)
                                {
                                    for (int i = 0; i < dtCompList.Rows.Count; i++)
                                    {
                                        Sheet CtrlSheet = workbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name.ToString().ToUpper() == Convert.ToString(dtCompList.Rows[i]["comp_desc"]).ToUpper().Trim()).FirstOrDefault();
                                        if (CtrlSheet != null)
                                        {
                                            int IntColCount = 1, IntCount = 0, IntRowsCount = 0;
                                            while (IntCount < 5)
                                            {
                                                IntColCount += 1;
                                                string StrSLNo = ClsExcel.GetCellValue(workbookPart, CtrlSheet, "A" + IntColCount);
                                                string StrAL = ClsExcel.GetCellValue(workbookPart, CtrlSheet, "B" + IntColCount);
                                                string StrReq = ClsExcel.GetCellValue(workbookPart, CtrlSheet, "C" + IntColCount);
                                                string StrKS = ClsExcel.GetCellValue(workbookPart, CtrlSheet, "E" + IntColCount);
                                                if (StrSLNo.Trim() != "" || StrAL.Trim() != "" || StrReq.Trim() != "" || StrKS.Trim() != "")
                                                {
                                                    IntCount = 0;
                                                    IntRowsCount += 1;
                                                    DataRow drReq = dtReqList.NewRow();
                                                    drReq["main_id"] = Convert.ToString(dtCompList.Rows[i]["ID"]).Trim();
                                                    drReq["req_comp_desc"] = Convert.ToString(dtCompList.Rows[i]["comp_desc"]).Trim();
                                                    drReq["sl_no"] = StrSLNo;
                                                    drReq["applicable_levels"] = StrAL;
                                                    drReq["requirements"] = StrReq;
                                                    drReq["knowledge_scale"] = StrKS;
                                                    dtReqList.Rows.Add(drReq);
                                                    dtReqList.AcceptChanges();
                                                    if (CompSubCheckRequiredField(StrSLNo, StrAL, StrReq, StrKS, Convert.ToString(dtCompList.Rows[i]["comp_desc"]).Trim()) == false)
                                                    {
                                                        StrError += Convert.ToString(Session["SInfoMsg"]).Trim();
                                                        Session["SInfoMsg"] = "";
                                                        Session["SInfoMsgType"] = "";
                                                        BlnResult = false;
                                                        IntError = 1;
                                                    }
                                                }
                                                else
                                                {
                                                    IntCount += 1;
                                                }
                                            }
                                            if (IntRowsCount == 0 )
                                            {
                                                StrError += CommonCls.GetErrorMsg("Competency description details not found in &quot;" + Convert.ToString(dtCompList.Rows[i]["comp_desc"]).Trim() + "&quot; sheet.");
                                                BlnResult = false;
                                                IntError = 1;
                                            }
                                            Session["Req_List"] = dtReqList;
                                        }
                                    }
                                }
                            }
                            // Other sheet required validation End
                        }
                    }
                }
                if (StrError != "")
                {
                    Session["SInfoMsg"] = StrError;
                    Session["SInfoMsgType"] = "ERROR";
                }
                return BlnResult;
            }
            catch (Exception ex)
            {
                logger.write("HomeController.cs", "ERROR ON CheckRequiredField() - " + ex.Message);
                Session["SInfoMsg"] = ex.Message;
                Session["SInfoMsgType"] = "ERROR";
                return false;
            }
        }
        #endregion
        #region "Dynamic Set Initial Row (DataTable)"
        private void SetInitialRow()
        {
            try
            {
                // Sheet DataTable
                DataTable dtSheetList = new DataTable();
                dtSheetList.Columns.Add(new DataColumn("ID", typeof(Int32)));
                dtSheetList.Columns.Add(new DataColumn("Relationship_Id", typeof(string)));
                dtSheetList.Columns.Add(new DataColumn("Sheet_Id", typeof(string)));
                dtSheetList.Columns.Add(new DataColumn("Sheet_Name", typeof(string)));
                dtSheetList.Columns["Id"].AutoIncrement = true;
                dtSheetList.Columns["Id"].AutoIncrementSeed = 1;
                dtSheetList.Columns["Id"].AutoIncrementStep = 1;
                dtSheetList.AcceptChanges();
                Session["Sheet_List"] = dtSheetList;

                // Emp DataTable
                DataTable dtEmpList = new DataTable();
                dtEmpList.Columns.Add(new DataColumn("ID", typeof(Int32)));
                dtEmpList.Columns.Add(new DataColumn("emp_name", typeof(string)));
                dtEmpList.Columns.Add(new DataColumn("emp_code", typeof(string)));
                dtEmpList.Columns.Add(new DataColumn("emp_exp_yrs_total", typeof(string)));
                dtEmpList.Columns.Add(new DataColumn("emp_exp_yrs_titan", typeof(string)));
                dtEmpList.Columns.Add(new DataColumn("emp_exp_yrs_outside", typeof(string)));
                dtEmpList.Columns.Add(new DataColumn("emp_level", typeof(string)));
                dtEmpList.Columns.Add(new DataColumn("emp_department", typeof(string)));
                dtEmpList.Columns.Add(new DataColumn("emp_designation", typeof(string)));
                dtEmpList.Columns.Add(new DataColumn("emp_supervisor", typeof(string)));
                dtEmpList.Columns.Add(new DataColumn("emp_reviewer", typeof(string)));
                dtEmpList.Columns["Id"].AutoIncrement = true;
                dtEmpList.Columns["Id"].AutoIncrementSeed = 1;
                dtEmpList.Columns["Id"].AutoIncrementStep = 1;
                dtEmpList.AcceptChanges();
                Session["Emp_List"] = dtEmpList;

                // Comp DataTable
                DataTable dtCompList = new DataTable();
                dtCompList.Columns.Add(new DataColumn("ID", typeof(Int32)));
                dtCompList.Columns.Add(new DataColumn("comp_type", typeof(string)));
                dtCompList.Columns.Add(new DataColumn("comp_desc", typeof(string)));
                dtCompList.Columns.Add(new DataColumn("comp_acl", typeof(string)));
                dtCompList.Columns.Add(new DataColumn("comp_rcl", typeof(string)));
                dtCompList.Columns["Id"].AutoIncrement = true;
                dtCompList.Columns["Id"].AutoIncrementSeed = 1;
                dtCompList.Columns["Id"].AutoIncrementStep = 1;
                dtCompList.AcceptChanges();
                Session["Comp_List"] = dtCompList;

                // Req DataTable
                DataTable dtReqList = new DataTable();
                dtReqList.Columns.Add(new DataColumn("ID", typeof(Int32)));
                dtReqList.Columns.Add(new DataColumn("main_id", typeof(string)));
                dtReqList.Columns.Add(new DataColumn("req_comp_desc", typeof(string)));
                dtReqList.Columns.Add(new DataColumn("sl_no", typeof(string)));
                dtReqList.Columns.Add(new DataColumn("applicable_levels", typeof(string)));
                dtReqList.Columns.Add(new DataColumn("requirements", typeof(string)));
                dtReqList.Columns.Add(new DataColumn("knowledge_scale", typeof(string)));
                dtReqList.Columns["Id"].AutoIncrement = true;
                dtReqList.Columns["Id"].AutoIncrementSeed = 1;
                dtReqList.Columns["Id"].AutoIncrementStep = 1;
                dtReqList.AcceptChanges();
                Session["Req_List"] = dtReqList;

            }
            catch (Exception ex)
            {
                logger.write("HomeController.cs", "ERROR ON SetInitialRow() - " + ex.Message);
            }
        }
        #endregion
        #region "Required Functions"
        private bool CheckRequiredField(General ObjValue)
        {
            Boolean BlnResult = true;
            string StrError = "";
            try
            {
                if (ObjValue.competancy_file == null || ObjValue.competancy_file.FileName == "")
                {
                    StrError += CommonCls.GetErrorMsg("Upload competency data sheet is required field.");
                    BlnResult = false;
                }
                if (StrError != "")
                {
                    Session["SInfoMsg"] = StrError;
                    Session["SInfoMsgType"] = "ERROR";
                }
                return BlnResult;
            }
            catch (Exception ex)
            {
                logger.write("HomeController.cs", "ERROR ON CheckRequiredField() - " + ex.Message);
                Session["SInfoMsg"] = ex.Message;
                Session["SInfoMsgType"] = "ERROR";
                return false;
            }
        }
        private bool EmpCheckRequiredField(string StrEmpName, string StrEmpCode, string StrEmpExpYrsTot, string StrEmpExpYrsTit, string StrEmpExpYrsOutSide, string StrEmpLevel, string StrEmpDept, string StrEmpDes, string StrEmpSup, string StrEmpReviewer)
        {
            Boolean BlnResult = true;
            string StrError = "";
            try
            {
                if (StrEmpName.Trim() == "")
                {
                    StrError += CommonCls.GetRequiredMsg("Employee name");
                    BlnResult = false;
                }
                if (StrEmpCode.Trim() == "")
                {
                    StrError += CommonCls.GetRequiredMsg("Employee code");
                    BlnResult = false;
                }
                if (StrEmpExpYrsTot.Trim() == "")
                {
                    StrError += CommonCls.GetRequiredMsg("Exp. In Yrs(Total)");
                    BlnResult = false;
                }
                if (StrEmpExpYrsTit.Trim() == "")
                {
                    StrError += CommonCls.GetRequiredMsg("Exp. In Yrs(Titan)");
                    BlnResult = false;
                }
                if (StrEmpExpYrsOutSide.Trim() == "")
                {
                    StrError += CommonCls.GetRequiredMsg("Exp. In Yrs(Outside)");
                    BlnResult = false;
                }
                if (StrEmpLevel.Trim() == "")
                {
                    StrError += CommonCls.GetRequiredMsg("Level");
                    BlnResult = false;
                }
                if (StrEmpDept.Trim() == "")
                {
                    StrError += CommonCls.GetRequiredMsg("Department");
                    BlnResult = false;
                }
                if (StrEmpDes.Trim() == "")
                {
                    StrError += CommonCls.GetRequiredMsg("Designation");
                    BlnResult = false;
                }
                if (StrEmpSup.Trim() == "")
                {
                    StrError += CommonCls.GetRequiredMsg("Supervisor");
                    BlnResult = false;
                }
                if (StrEmpReviewer.Trim() == "")
                {
                    StrError += CommonCls.GetRequiredMsg("Reviewer");
                    BlnResult = false;
                }
                if (StrError != "")
                {
                    Session["SInfoMsg"] = StrError;
                    Session["SInfoMsgType"] = "ERROR";
                }
                return BlnResult;
            }
            catch (Exception ex)
            {
                logger.write("HomeController.cs", "ERROR ON EmpCheckRequiredField() - " + ex.Message);
                Session["SInfoMsg"] = ex.Message;
                Session["SInfoMsgType"] = "ERROR";
                return false;
            }
        }
        private bool CompCheckRequiredField(string StrCompType, string StrCompDesc, string StrACL, string StrRCL)
        {
            Boolean BlnResult = true;
            string StrError = "";
            try
            {
                if (StrCompType.Trim() == "")
                {
                    StrError += CommonCls.GetRequiredMsg("Competency Type");
                    BlnResult = false;
                }
                if (StrCompDesc.Trim() == "")
                {
                    StrError += CommonCls.GetRequiredMsg("Competency description");
                    BlnResult = false;
                }
                if (StrACL.Trim() == "")
                {
                    StrError += CommonCls.GetRequiredMsg("Actual competency level");
                    BlnResult = false;
                }
                if (StrRCL.Trim() == "")
                {
                    StrError += CommonCls.GetRequiredMsg("Required competency level");
                    BlnResult = false;
                }
                if (StrError != "")
                {
                    Session["SInfoMsg"] = StrError;
                    Session["SInfoMsgType"] = "ERROR";
                }
                return BlnResult;
            }
            catch (Exception ex)
            {
                logger.write("HomeController.cs", "ERROR ON EmpCheckRequiredField() - " + ex.Message);
                Session["SInfoMsg"] = ex.Message;
                Session["SInfoMsgType"] = "ERROR";
                return false;
            }
        }
        private bool CompSubCheckRequiredField(string StrSLNo, string StrAL, string StrReq, string StrKS, string StrSheetName)
        {
            Boolean BlnResult = true;
            string StrError = "";
            try
            {
                if (StrSLNo.Trim() == "")
                {
                    StrError += CommonCls.GetRequiredMsg(StrSheetName + " - Sl No");
                    BlnResult = false;
                }
                if (StrAL.Trim() == "")
                {
                    StrError += CommonCls.GetRequiredMsg(StrSheetName + " - Applicable levels");
                    BlnResult = false;
                }
                if (StrReq.Trim() == "")
                {
                    StrError += CommonCls.GetRequiredMsg(StrSheetName + " - Requirements");
                    BlnResult = false;
                }
                if (StrKS.Trim() == "")
                {
                    StrError += CommonCls.GetRequiredMsg(StrSheetName + " - Knowledge scale");
                    BlnResult = false;
                }
                if (StrError != "")
                {
                    Session["SInfoMsg"] = StrError;
                    Session["SInfoMsgType"] = "ERROR";
                }
                return BlnResult;
            }
            catch (Exception ex)
            {
                logger.write("HomeController.cs", "ERROR ON CompSubCheckRequiredField() - " + ex.Message);
                Session["SInfoMsg"] = ex.Message;
                Session["SInfoMsgType"] = "ERROR";
                return false;
            }
        }
        #endregion
        #region "View Data"
        public ActionResult DBView()
        {
            if (Convert.ToString(Session["SInfoMsg"]) != "")
            {
                CtrlGen.Message_Info = CommonCls.GetMsg(Convert.ToString(Session["SInfoMsg"]).Trim(), Convert.ToString(Session["SInfoMsgType"]).Trim());
                Session["SInfoMsg"] = "";
                Session["SInfoMsgType"] = "";
            }
            CtrlGen.table_data = Get_Table_Data("comp_list");
            return View(CtrlGen);
        }
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult DBView(General ObjValue)
        {
            try
            {
                ObjValue.table_data = Get_Table_Data(ObjValue.table_name);
            }
            catch (Exception ex)
            {
                logger.write("HomeController.cs", "ERROR ON DBView(ObjValue) - " + ex.Message);
                ObjValue.Message_Info = CommonCls.GetMsg(ex.Message, "ERROR");
            }
            return View(ObjValue);
        }
        private string Get_Table_Data(string StrTableName)
        {
            string StrHTML = "";
            try
            {
                DataTable dt = new DataTable();
                dt = DBCon.GetDataTableFromDb("SELECT * FROM " + StrTableName + " ORDER BY 1");
                if(dt != null)
                {
                    List<string> ArrayColumns = new List<string>();
                    foreach (DataColumn CtrlColumn in dt.Columns)
                    {
                        ArrayColumns.Add(CtrlColumn.ColumnName);
                    }
                    StrHTML += "<table id='MyTable' class='table table-striped table-bordered' style='width:100% !important;'>";
                    StrHTML += "<thead><tr>";
                    for (int i = 0; i < ArrayColumns.Count; i++)
                    {
                        StrHTML += "<th>" + Convert.ToString(ArrayColumns[i]).Trim()  + "</th>";
                    }
                    StrHTML += "</tr></thead>";
                    if(dt.Rows.Count > 0)
                    {
                        StrHTML += "<tbody>";
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            StrHTML += "<tr>";
                            for (int j = 0; j < ArrayColumns.Count; j++)
                            {
                                StrHTML += "<td>" + Convert.ToString(dt.Rows[i][Convert.ToString(ArrayColumns[j])]).Trim().Replace("\"", "&quot;").Replace("'", "&apos;").Replace("`", "&#96;").Replace("＂", "&#65282;").Replace("＇", "&#65287;") + "</td>";
                            }
                            StrHTML += "</tr>";
                        }
                        StrHTML += "</tbody>";
                    }
                    StrHTML += "</table>";
                }
            }
            catch (Exception ex)
            {
                logger.write("HomeController.cs", "ERROR ON Get_Table_Data() - " + ex.Message);
                StrHTML ="";
            }
            return StrHTML;
        }
        #endregion
    }
}