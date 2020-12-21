//using System.Xml.Serialization;

//namespace AIT.TFS.SyncService.Service.Configuration.Serialization.TestReport
//{
//    //[XmlRoot("TestSpecReportDefaults")]
//    public class TestSpecReportingDefault //: ITestSpecReportDefault
//    {
//        [XmlElement("CreateDocumentStructure")]
//        public string CreateDocumentStructure { get; set; }

//        [XmlElement("IncludeTestConfigurations")]
//        public string TestSpecReport_ConfigurationPositionType { get; set; }

//        [XmlElement("SortTestCasesBy")]
//        public string TestSpecReport_TestCaseSortType { get; set; }

//        //internal void ReplaceValuesIfNotNull(ITestSpecReportDefault value)
//        //{
//        //    var newValue = value.TestSpecReport_CreateDocumentStructure;
//        //    var stayValue = TestSpecReport_CreateDocumentStructure;
//        //    TestSpecReport_CreateDocumentStructure = !string.IsNullOrEmpty(newValue) ? newValue : stayValue;

//        //    newValue = value.TestSpecReport_ConfigurationPositionType;
//        //    stayValue = TestSpecReport_ConfigurationPositionType;
//        //    TestSpecReport_ConfigurationPositionType = !string.IsNullOrEmpty(newValue) ? newValue : stayValue;

//        //    newValue = value.TestSpecReport_TestCaseSortType;
//        //    stayValue = TestSpecReport_TestCaseSortType;
//        //    TestSpecReport_TestCaseSortType = !string.IsNullOrEmpty(newValue) ? newValue : stayValue;
//        //}
//    }
//}