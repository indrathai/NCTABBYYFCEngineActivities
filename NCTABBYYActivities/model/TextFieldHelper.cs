using FCEngine;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NCTABBYYActivities.model
{
    /// <summary>
    /// Helping class for TextFieldPresenter.
    /// </summary>
    internal class TextFieldHelper
    {
        private static int ALL_CHAR_PARAMS =
            (int)(FCEngine.CharParamsFlags.CFL_EditedByRules |
            FCEngine.CharParamsFlags.CFL_EditedByUser |
            FCEngine.CharParamsFlags.CFL_LanguageId |
            FCEngine.CharParamsFlags.CFL_LanguageName |
            FCEngine.CharParamsFlags.CFL_NeedsGroupVerification |
            FCEngine.CharParamsFlags.CFL_NeedsVerification |
            FCEngine.CharParamsFlags.CFL_PageID |
            FCEngine.CharParamsFlags.CFL_Rectangle |
            FCEngine.CharParamsFlags.CFL_Suspicious);

        /// <summary>
        /// Creates new copy of Flexi Capture text object.
        /// </summary>
        /// <param name="sourceText">Source text object from getting a copy.</param>
        /// <param name="engine">FCEngine.IEngine object.</param>
        /// <returns>New Flexi Capture text object.</returns>
        public static FCEngine.IText CreateNewCopyFrom(FCEngine.IText sourceText, FCEngine.IEngine engine)
        {
            FCEngine.IText newText = engine.CreateText(sourceText.Text, null);

            FCEngine.ICharParams[] oldCharParams = GetArrayOfCharParams(sourceText, engine);
            for (int i = 0; i < oldCharParams.Length; i++)
            {
                newText.SetCharParams(i, 1, oldCharParams[i], ALL_CHAR_PARAMS);
            }
            return newText;
        }

        /// <summary>
        /// Returns true if text has unverified characters.
        /// </summary>
        /// <param name="text">Flexi Capture text object.</param>
        /// <param name="engine">FCEngine.IEngine object.</param>
        /// <returns>Boolean value.</returns>
        public static bool TextHasUnverifiedCharacters(FCEngine.IText text, FCEngine.IEngine engine)
        {
            for (int i = 0; i < text.Length; i++)
            {
                FCEngine.ICharParams charParams = engine.CreateCharParams();
                text.GetCharParams(i, charParams);
                if (charParams.NeedsVerification)
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// Compares two text objects for equality.
        /// </summary>
        /// <param name="text1">First Flexi Capture text object.</param>
        /// <param name="text2">Second Flexi Capture text object.</param>
        /// <param name="engine">FCEngine.IEngine object.</param>
        /// <returns>Boolean value.</returns>
        public static bool TextObjectsAreSame(FCEngine.IText text1, FCEngine.IText text2, FCEngine.IEngine engine)
        {
            if (text1.Length != text2.Length)
            {
                return false;
            }

            FCEngine.ICharParams charParams = engine.CreateCharParams();
            for (int i = 0; i < text1.Length; i++)
            {
                if (text1.Text[i] != text2.Text[i] ||
                    characterNeedsVerification(text1, i, charParams) != characterNeedsVerification(text2, i, charParams))
                {
                    return false;
                }
            }

            return true;
        }

        /// <summary>
        /// Returns array of char params objects.
        /// </summary>
        /// <param name="text">Flexi Capture text object for getting char params.</param>
        /// <param name="engine">FCEngine.IEngine object.</param>
        /// <returns>Array of char params.</returns>
        public static FCEngine.ICharParams[] GetArrayOfCharParams(FCEngine.IText text, FCEngine.IEngine engine)
        {
            List<FCEngine.ICharParams> charParamsList = new List<FCEngine.ICharParams>();

            for (int i = 0; i < text.Length; i++)
            {
                FCEngine.ICharParams charParams = engine.CreateCharParams();
                text.GetCharParams(i, charParams);
                charParamsList.Add(charParams);
            }

            return charParamsList.ToArray();
        }

        private static bool characterNeedsVerification(FCEngine.IText text, int characterNumber, FCEngine.ICharParams charParams)
        {
            text.GetCharParams(characterNumber, charParams);
            return charParams.NeedsVerification;
        }

        public static int GetConfidenLevel(FCEngine.IEngine engine, FCEngine.IDocument document)
        {
            int result = 0;
            int charAll = 0;
            int charNeedsVerification = 0;
            IField root = document as IField;
            var fieldsAll = recursiveFindFieldLast(root);
            if (!object.ReferenceEquals(fieldsAll, null))
            {
                for (int i = 0; i < fieldsAll.Count; i++)
                {
                    var type = fieldsAll[i].Type;
                    if (type == FieldTypeEnum.FT_TextField || type == FieldTypeEnum.FT_DateTimeField || type == FieldTypeEnum.FT_CurrencyField || type == FieldTypeEnum.FT_NumberField)
                    {
                        if (!object.ReferenceEquals(fieldsAll[i].Value, null))
                        {
                            FCEngine.IText text = fieldsAll[i].Value.AsText;
                            for (int d = 0; d < text.Length; d++)
                            {
                                FCEngine.ICharParams charParams = engine.CreateCharParams();
                                text.GetCharParams(d, charParams);
                                if (charParams.NeedsVerification)
                                    charNeedsVerification += 1;
                            }
                            charAll += text.Length;
                        }
                        else if (!object.ReferenceEquals(fieldsAll[i].Instances, null))
                        {
                            var instances = fieldsAll[i].Instances;
                            for (int o = 0; o < instances.Count; o++)
                            {
                                if (!object.ReferenceEquals(instances[o].Value, null))
                                {
                                    FCEngine.IText text = instances[o].Value.AsText;
                                    for (int d = 0; d < text.Length; d++)
                                    {
                                        FCEngine.ICharParams charParams = engine.CreateCharParams();
                                        text.GetCharParams(d, charParams);
                                        if (charParams.NeedsVerification)
                                            charNeedsVerification += 1;

                                    }
                                    charAll += text.Length;
                                }
                            }
                        }

                    }
                    else if (type == FieldTypeEnum.FT_Table)
                    {
                        if (!object.ReferenceEquals(fieldsAll[i].Instances, null))
                        {
                            var instances = fieldsAll[i].Instances;
                            for (int o = 0; o < instances.Count; o++)
                            {
                                var row = instances.Item(o);
                                var cells = row.Children;
                                if (!object.ReferenceEquals(cells, null))
                                {
                                    for (int j = 0; j < cells.Count; j++)
                                    {
                                        var cellField = cells.Item(j);
                                        var cellType = cellField.Type;
                                        if (cellType == FieldTypeEnum.FT_TextField || cellType == FieldTypeEnum.FT_DateTimeField || cellType == FieldTypeEnum.FT_CurrencyField || cellType == FieldTypeEnum.FT_NumberField)
                                        {
                                            FCEngine.IText text = cellField.Value.AsText;
                                            for (int d = 0; d < text.Length; d++)
                                            {
                                                FCEngine.ICharParams charParams = engine.CreateCharParams();
                                                text.GetCharParams(d, charParams);
                                                if (charParams.NeedsVerification)
                                                    charNeedsVerification += 1;

                                            }
                                            charAll += text.Length;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                if (charAll != 0)
                    result = 100 - ((charNeedsVerification * 100) / charAll);
            }
            return result;
        }

        public static bool isValidEightItems(FCEngine.IEngine engine, FCEngine.IDocument document, string[] items)
        {
            IField root = document as IField;
            var fieldsAll = recursiveFindFieldLast(root);
            if (!object.ReferenceEquals(fieldsAll, null))
            {
                for (int i = 0; i < fieldsAll.Count; i++)
                {
                    var type = fieldsAll[i].Type;
                    var fieldName = fieldsAll[i].Name;
                    if (Array.IndexOf(items, fieldName) >= 0)
                    {

                        if (type == FieldTypeEnum.FT_TextField || type == FieldTypeEnum.FT_DateTimeField || type == FieldTypeEnum.FT_CurrencyField || type == FieldTypeEnum.FT_NumberField)
                        {
                            if (!object.ReferenceEquals(fieldsAll[i].Value, null))
                            {
                                FCEngine.IText text = fieldsAll[i].Value.AsText;

                                Console.WriteLine("fieldName: " + fieldName + " and value: " + text);
                                if (!Object.ReferenceEquals(text, null))
                                {
                                    if (Object.ReferenceEquals(text.PlainText, "") || text.PlainText.Length == 0)
                                    {
                                        return false;
                                    }
                                }
                            }
                        }
                    }
                    else if (type == FieldTypeEnum.FT_Table)
                    {
                        if (!object.ReferenceEquals(fieldsAll[i].Instances, null))
                        {
                            var instances = fieldsAll[i].Instances;
                            for (int o = 0; o < instances.Count; o++)
                            {
                                var row = instances.Item(o);
                                var cells = row.Children;
                                if (!object.ReferenceEquals(cells, null))
                                {
                                    for (int j = 0; j < cells.Count; j++)
                                    {
                                        var cellField = cells.Item(j);
                                        var cellType = cellField.Type;
                                        var cellname = cellField.Name;
                                        if (Array.IndexOf(items, cellname) >= 0)
                                        {
                                            if (cellType == FieldTypeEnum.FT_TextField || cellType == FieldTypeEnum.FT_DateTimeField || cellType == FieldTypeEnum.FT_CurrencyField || cellType == FieldTypeEnum.FT_NumberField)
                                            {
                                                FCEngine.IText text = cellField.Value.AsText;
                                                Console.WriteLine("cellName: " + cellname + " and value: " + text);
                                                if (!Object.ReferenceEquals(text, null))
                                                {
                                                    if (Object.ReferenceEquals(text.PlainText, "") || text.PlainText.Length == 0)
                                                    {
                                                        return false;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            return true;
        }

        public static int GetAccurracyLevelHeaderMandatoryField(FCEngine.IEngine engine, FCEngine.IDocument document, string[] mandatoryFields)
        {
            int result = 0;
            int charAll = 0;
            int charNeedsVerification = 0;
            IField root = document as IField;
            var fieldsAll = recursiveFindFieldLast(root);

            /*dataTable = new DataTable();

            //setup table
            dataTable.Columns.Add("Field Name", typeof(string));
            dataTable.Columns.Add("Original Value", typeof(string));
            dataTable.Columns.Add("Idx Need to Verify", typeof(string));
            dataTable.Columns.Add("Need to Verify", typeof(string));
            */

            if (!object.ReferenceEquals(fieldsAll, null))
            {
                for (int i = 0; i < fieldsAll.Count; i++)
                {
                    int fcharAll = 0;
                    int fcharNeedsVerification = 0;
                    var type = fieldsAll[i].Type;
                    var fieldName = fieldsAll[i].Name;
                    var fieldValue = "";
                    string[] fieldVerification = null;
                    string strNeedVerification = "";
                    //calculate only for Mandatory field
                    if (Array.IndexOf(mandatoryFields, fieldName) >= 0)
                    {
                        if (type == FieldTypeEnum.FT_TextField || type == FieldTypeEnum.FT_DateTimeField || type == FieldTypeEnum.FT_CurrencyField || type == FieldTypeEnum.FT_NumberField)
                        {
                            if (!object.ReferenceEquals(fieldsAll[i].Value, null))
                            {
                                FCEngine.IText text = fieldsAll[i].Value.AsText;
                                fieldValue = text.PlainText;
                                fieldVerification = new string[fieldValue.Length];
                                for (int d = 0; d < text.Length; d++)
                                {
                                    FCEngine.ICharParams charParams = engine.CreateCharParams();
                                    text.GetCharParams(d, charParams);
                                    fieldVerification[d] = "";
                                    if (charParams.NeedsVerification)
                                    {
                                        charNeedsVerification += 1;
                                        fcharNeedsVerification += 1;
                                        strNeedVerification = strNeedVerification + fieldValue.ToCharArray(d, 1)[0].ToString();
                                    }
                                }
                                charAll += text.Length;
                                fcharAll += text.Length;
                            }
                        }

                        Console.WriteLine(string.Format("field of {0} has total char {1} and value {2} then total char need to verify {3} and value {4})", fieldName, fcharAll, fieldValue, fcharNeedsVerification, strNeedVerification));
                    }
                }
                if (charAll != 0)
                    result = 100 - ((charNeedsVerification * 100) / charAll);
            }

            Console.WriteLine(string.Format("Confidence Header Level is {0} ", result.ToString()));
            return result;
        }

        public static string GetAccurracyLevelHeaderMandatoryFieldWithTotalCharacters(FCEngine.IEngine engine, FCEngine.IDocument document, string[] mandatoryFields)
        {
            int result = 0;
            int charAll = 0;
            int charNeedsVerification = 0;
            IField root = document as IField;
            var fieldsAll = recursiveFindFieldLast(root);

            /*dataTable = new DataTable();

            //setup table
            dataTable.Columns.Add("Field Name", typeof(string));
            dataTable.Columns.Add("Original Value", typeof(string));
            dataTable.Columns.Add("Idx Need to Verify", typeof(string));
            dataTable.Columns.Add("Need to Verify", typeof(string));
            */

            if (!object.ReferenceEquals(fieldsAll, null))
            {
                for (int i = 0; i < fieldsAll.Count; i++)
                {
                    int fcharAll = 0;
                    int fcharNeedsVerification = 0;
                    var type = fieldsAll[i].Type;
                    var fieldName = fieldsAll[i].Name;
                    var fieldValue = "";
                    string[] fieldVerification = null;
                    string strNeedVerification = "";

                    //calculate only for Mandatory field
                    if (Array.IndexOf(mandatoryFields, fieldName) >= 0)
                    {
                        if (type == FieldTypeEnum.FT_TextField || type == FieldTypeEnum.FT_DateTimeField || type == FieldTypeEnum.FT_CurrencyField || type == FieldTypeEnum.FT_NumberField)
                        {
                            if (!object.ReferenceEquals(fieldsAll[i].Value, null))
                            {
                                FCEngine.IText text = fieldsAll[i].Value.AsText;
                                fieldValue = text.PlainText;
                                Console.WriteLine(string.Format("plain text: {0}", fieldValue));
                                fieldVerification = new string[fieldValue.Length];
                                for (int d = 0; d < text.Length; d++)
                                {
                                    FCEngine.ICharParams charParams = engine.CreateCharParams();
                                    text.GetCharParams(d, charParams);
                                    fieldVerification[d] = "";
                                    if (charParams.NeedsVerification)
                                    {
                                        charNeedsVerification += 1;
                                        fcharNeedsVerification += 1;
                                        strNeedVerification = strNeedVerification + fieldValue.ToCharArray(d, 1)[0].ToString();
                                    }
                                }
                                charAll += text.Length;
                                fcharAll += text.Length;
                            }
                        }

                        Console.WriteLine(string.Format("field of {0} has total char {1} and value {2} then total char need to verify {3} and value {4})", fieldName, fcharAll, fieldValue, fcharNeedsVerification, strNeedVerification));
                    }
                }
                if (charAll != 0)
                    result = 100 - ((charNeedsVerification * 100) / charAll);
            }

            Console.WriteLine(string.Format("Confidence Header Level is {0} ", result.ToString()));
            return string.Format("{0} % ( {1} of {2} )", result, charAll - charNeedsVerification, charAll);
        }

        public static int GetAccurracyLevelDetailMandatoryField(FCEngine.IEngine engine, FCEngine.IDocument document, string[] mandatoryFields)
        {
            int result = 0;
            int charAll = 0;
            int charNeedsVerification = 0;
            IField root = document as IField;
            var fieldsAll = recursiveFindFieldLast(root);


            if (!object.ReferenceEquals(fieldsAll, null))
            {
                for (int i = 0; i < fieldsAll.Count; i++)
                {

                    var type = fieldsAll[i].Type;
                    var fieldName = fieldsAll[i].Name;
                    var fieldValue = "";
                    string[] fieldVerification = null;

                    //calculate only for Mandatory field
                    if (type == FieldTypeEnum.FT_Table)
                    {
                        if (!object.ReferenceEquals(fieldsAll[i].Instances, null))
                        {
                            var instances = fieldsAll[i].Instances;
                            for (int o = 0; o < instances.Count; o++)
                            {
                                var row = instances.Item(o);
                                var cells = row.Children;
                                if (!object.ReferenceEquals(cells, null))
                                {
                                    for (int j = 0; j < cells.Count; j++)
                                    {
                                        var cellField = cells.Item(j);
                                        var cellType = cellField.Type;
                                        var cellName = cellField.Name;
                                        int fcharAll = 0;
                                        int fcharNeedsVerification = 0;
                                        string strNeedVerification = "";
                                        if (Array.IndexOf(mandatoryFields, cellName) >= 0)
                                        {
                                            if (cellType == FieldTypeEnum.FT_TextField || cellType == FieldTypeEnum.FT_DateTimeField || cellType == FieldTypeEnum.FT_CurrencyField || cellType == FieldTypeEnum.FT_NumberField)
                                            {
                                                FCEngine.IText text = cellField.Value.AsText;
                                                fieldValue = text.PlainText;
                                                fieldVerification = new string[fieldValue.Length];
                                                for (int d = 0; d < text.Length; d++)
                                                {
                                                    FCEngine.ICharParams charParams = engine.CreateCharParams();
                                                    text.GetCharParams(d, charParams);
                                                    if (charParams.NeedsVerification)
                                                    {
                                                        charNeedsVerification += 1;
                                                        fcharNeedsVerification += 1;
                                                        strNeedVerification = strNeedVerification + fieldValue.ToCharArray(d, 1)[0].ToString();
                                                        fieldVerification[d] = fieldValue.ToCharArray(d, 1)[0].ToString();
                                                    }
                                                }
                                                charAll += text.Length;
                                                fcharAll += text.Length;
                                            }
                                            Console.WriteLine(string.Format("field of {0} has total char {1} and value {2} then total char need to verify {3} and value {4})", cellName, fcharAll, fieldValue, fcharNeedsVerification, strNeedVerification));
                                        }
                                    }
                                }
                            }
                        }
                    }

                }
                if (charAll != 0)
                    result = 100 - ((charNeedsVerification * 100) / charAll);
            }

            Console.WriteLine(string.Format("Confidence Detail Level is {0} ", result.ToString()));
            return result;
        }

        public static string GetAccurracyLevelDetailMandatoryFieldWithTotalCharacter(FCEngine.IEngine engine, FCEngine.IDocument document, string[] mandatoryFields)
        {
            int result = 0;
            int charAll = 0;
            int charNeedsVerification = 0;
            IField root = document as IField;
            var fieldsAll = recursiveFindFieldLast(root);


            if (!object.ReferenceEquals(fieldsAll, null))
            {
                for (int i = 0; i < fieldsAll.Count; i++)
                {

                    var type = fieldsAll[i].Type;
                    var fieldName = fieldsAll[i].Name;
                    var fieldValue = "";
                    string[] fieldVerification = null;

                    //calculate only for Mandatory field
                    if (type == FieldTypeEnum.FT_Table)
                    {
                        if (!object.ReferenceEquals(fieldsAll[i].Instances, null))
                        {
                            var instances = fieldsAll[i].Instances;
                            for (int o = 0; o < instances.Count; o++)
                            {
                                var row = instances.Item(o);
                                var cells = row.Children;
                                if (!object.ReferenceEquals(cells, null))
                                {
                                    for (int j = 0; j < cells.Count; j++)
                                    {
                                        var cellField = cells.Item(j);
                                        var cellType = cellField.Type;
                                        var cellName = cellField.Name;
                                        int fcharAll = 0;
                                        int fcharNeedsVerification = 0;
                                        string strNeedVerification = "";
                                        if (Array.IndexOf(mandatoryFields, cellName) >= 0)
                                        {
                                            if (cellType == FieldTypeEnum.FT_TextField || cellType == FieldTypeEnum.FT_DateTimeField || cellType == FieldTypeEnum.FT_CurrencyField || cellType == FieldTypeEnum.FT_NumberField)
                                            {
                                                FCEngine.IText text = cellField.Value.AsText;
                                                fieldValue = text.PlainText;
                                                fieldVerification = new string[fieldValue.Length];
                                                for (int d = 0; d < text.Length; d++)
                                                {
                                                    FCEngine.ICharParams charParams = engine.CreateCharParams();
                                                    text.GetCharParams(d, charParams);
                                                    if (charParams.NeedsVerification)
                                                    {
                                                        charNeedsVerification += 1;
                                                        fcharNeedsVerification += 1;
                                                        strNeedVerification = strNeedVerification + fieldValue.ToCharArray(d, 1)[0].ToString();
                                                    }
                                                }
                                                charAll += text.Length;
                                                fcharAll += text.Length;
                                            }
                                            Console.WriteLine(string.Format("field of {0} has total char {1} and value {2} then total char need to verify {3} and value {4})", cellName, fcharAll, fieldValue, fcharNeedsVerification, strNeedVerification));
                                        }
                                    }
                                }
                            }
                        }
                    }

                }
                if (charAll != 0)
                    result = 100 - ((charNeedsVerification * 100) / charAll);
            }

            Console.WriteLine(string.Format("Accuracy Detail Level is {0} ", result.ToString()));
            return string.Format("{0} % ( {1} of {2} )", result, charAll - charNeedsVerification, charAll);
        }

        public static string GetConfidenLevelWithTotalCharacters(FCEngine.IEngine engine, FCEngine.IDocument document)
        {
            int result = 0;
            int charAll = 0;
            int charNeedsVerification = 0;
            IField root = document as IField;
            var fieldsAll = recursiveFindFieldLast(root);
            if (!object.ReferenceEquals(fieldsAll, null))
            {
                for (int i = 0; i < fieldsAll.Count; i++)
                {
                    var type = fieldsAll[i].Type;
                    if (type == FieldTypeEnum.FT_TextField || type == FieldTypeEnum.FT_DateTimeField
                        || type == FieldTypeEnum.FT_CurrencyField || type == FieldTypeEnum.FT_NumberField)
                    {
                        if (!object.ReferenceEquals(fieldsAll[i].Value, null))
                        {
                            FCEngine.IText text = fieldsAll[i].Value.AsText;
                            for (int d = 0; d < text.Length; d++)
                            {
                                FCEngine.ICharParams charParams = engine.CreateCharParams();
                                text.GetCharParams(d, charParams);
                                if (charParams.NeedsVerification)
                                    charNeedsVerification += 1;

                            }
                            charAll += text.Length;
                        }
                        else if (!object.ReferenceEquals(fieldsAll[i].Instances, null))
                        {
                            var instances = fieldsAll[i].Instances;
                            for (int o = 0; o < instances.Count; o++)
                            {
                                if (!object.ReferenceEquals(instances[o].Value, null))
                                {
                                    FCEngine.IText text = instances[o].Value.AsText;
                                    for (int d = 0; d < text.Length; d++)
                                    {
                                        FCEngine.ICharParams charParams = engine.CreateCharParams();
                                        text.GetCharParams(d, charParams);
                                        if (charParams.NeedsVerification)
                                            charNeedsVerification += 1;

                                    }
                                    charAll += text.Length;
                                }
                            }
                        }

                    }
                    else if (type == FieldTypeEnum.FT_Table)
                    {
                        if (!object.ReferenceEquals(fieldsAll[i].Instances, null))
                        {
                            var instances = fieldsAll[i].Instances;
                            for (int o = 0; o < instances.Count; o++)
                            {
                                var row = instances.Item(o);
                                var cells = row.Children;
                                if (!object.ReferenceEquals(cells, null))
                                {
                                    for (int j = 0; j < cells.Count; j++)
                                    {
                                        var cellField = cells.Item(j);
                                        var cellType = cellField.Type;
                                        if (cellType == FieldTypeEnum.FT_TextField || cellType == FieldTypeEnum.FT_DateTimeField || cellType == FieldTypeEnum.FT_CurrencyField || cellType == FieldTypeEnum.FT_NumberField)
                                        {
                                            FCEngine.IText text = cellField.Value.AsText;
                                            for (int d = 0; d < text.Length; d++)
                                            {
                                                FCEngine.ICharParams charParams = engine.CreateCharParams();
                                                text.GetCharParams(d, charParams);
                                                if (charParams.NeedsVerification)
                                                    charNeedsVerification += 1;

                                            }
                                            charAll += text.Length;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                if (charAll != 0)
                    result = 100 - ((charNeedsVerification * 100) / charAll);
            }
            return string.Format("{0} % ( {1} of {2} )", result, charAll - charNeedsVerification, charAll);
        }

        static System.Collections.Generic.List<IField> recursiveFindFieldLast(IField node)
        {
            System.Collections.Generic.List<IField> result = null;
            if (node != null)
            {
                IFields children = node.Children;
                if (children != null)
                {
                    result = new System.Collections.Generic.List<IField>();
                    foreach (IField child in children)
                    {
                        result.Add(child);

                        System.Collections.Generic.List<IField> founds = recursiveFindFieldLast(child);
                        if (founds != null && founds.Count > 0)
                            result.AddRange(founds);
                    }
                }
            }
            return result;
        }

        public static void Using_recognition_variants_and_extended_character_info(IEngine engine, IDocument document, string fieldName)
        {
            Console.WriteLine("Enable recognition variants...");
            engine.EnableRecognitionVariants(true);

            Console.WriteLine("Find the field of interest...");
            IField field = findField(document, fieldName);

            Console.WriteLine("Show word variants...");
            // (1) You might want to use word variants if you know how to choose the correct result
            // and this knowledge cannot be communicated to the engine (in the form of a regular expression 
            // or dictionary) or if you decide that you can do the selection better.
            // EXAMPLES: (1) a field containing some code with a checksum (2) a field that can be looked up 
            // in a database (3) a field that can be crosschecked with another field in the same document
            IText text = field.Value.AsText;
            IRecognizedWordInfo wordInfo = engine.CreateRecognizedWordInfo();
            // Get the recognition info for the first word. Thus we'll know the number of available variants
            text.GetRecognizedWord(0, -1, wordInfo);
            for (int i = 0; i < wordInfo.RecognitionVariantsCount; i++)
            {
                // Get the specified recognition variant for the word
                text.GetRecognizedWord(0, i, wordInfo);
                Console.WriteLine(wordInfo.Text);
            }
            Console.WriteLine("Ended ...");

            Console.WriteLine("Show char variants for each word variant...");
            // (2) You can use a more advanced approach and build your own hypotheses from variants
            // for individual characters. This approach can be very computationally intensive, so use it with care.
            // Use word/char variant confidence to limit the number of built hypotheses
            IRecognizedCharacterInfo charInfo = engine.CreateRecognizedCharacterInfo();
            for (int k = 0; k < wordInfo.RecognitionVariantsCount; k++)
            {
                // For each variant of the first word
                text.GetRecognizedWord(0, k, wordInfo);
                Console.WriteLine(wordInfo.Text);
                for (int i = 0; i < wordInfo.Text.Length; i++)
                {
                    // Get the recognition info for the first character (the number of available variants)
                    wordInfo.GetRecognizedCharacter(i, -1, charInfo);
                    string charVars = "";
                    for (int j = 0; j < charInfo.RecognitionVariantsCount; j++)
                    {
                        // Get the specified recognition variant for the character. The variant may contain
                        // more than one character if the geometry in the specified position can be interpreted
                        // as several merged characters. For example, something which looks like poorly printed 'U' 
                        // can actually be a pair of merged 'I'-s or 'I' + 'J'
                        wordInfo.GetRecognizedCharacter(i, j, charInfo);
                        if (charInfo.CharConfidence > 50)
                        {
                            charVars += charInfo.Character.PadRight(4, ' ');
                        }
                    }
                    Console.WriteLine(charVars);
                }
                Console.WriteLine("Ended");
            }
            Console.WriteLine("Ended");

            Console.WriteLine("Linking text in the field to recognition variants...");
            // (3) You can find corresponding recognition word and character variants for each character in the text
            // even if the text has been modified. This can be helpful in building your own advanced verification tools
            // where users can choose variants for words while typing or from a list
            ICharParams charParams = engine.CreateCharParams();
            for (int i = 0; i < text.Length; i++)
            {
                text.GetCharParams(i, charParams);
                Console.WriteLine(string.Format("'{0}' is char number {1} in word number {2}", text.Text[i],
                    charParams.RecognizedWordCharacterIndex, charParams.RecognizedWordIndex));
            }
            Console.WriteLine("");

            Console.WriteLine("Obtaining extended char information for characters in text...");
            // (4) You can obtain extended char info for a character in the text without getting involved deeply in  
            // the recognition variants API by means of the shortcut method GetCharParamsEx
            for (int i = 0; i < text.Length; i++)
            {
                text.GetCharParamsEx(i, charParams, charInfo);
                Console.WriteLine(string.Format("'{0}' serif probability is {1}", charInfo.Character, charInfo.SerifProbability));
            }
            Console.WriteLine("");

            // Restore the initial state of the engine. It is optional and should not be done if you always
            // use recognition variants. Doing so will reset some internal caches and if done repeatedly 
            // might affect performance
            engine.EnableRecognitionVariants(false);
        }

        static public IField findField(IDocument document, string name)
        {
            IField root = document as IField;
            return recursiveFindField(root, name);
        }

        static IField recursiveFindField(IField node, string name)
        {
            IFields children = node.Children;
            if (children != null)
            {
                foreach (IField child in children)
                {
                    if (child.Name == name)
                    {
                        return child;
                    }
                    else
                    {
                        IField found = recursiveFindField(child, name);
                        if (found != null)
                        {
                            return found;
                        }
                    }
                }
            }
            return null;
        }
    }
}
