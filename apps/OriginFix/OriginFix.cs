using System;
using System.Runtime.InteropServices;
using Kompas6API5;
using Kompas6Constants;
using Kompas6Constants3D;
using Base;


namespace RunCommands
{
    /// <summary>
    /// Create new part with english namings
    /// </summary>
    ///
    // ReSharper disable once UnusedType.Global
    public class OriginFix : BaseCommand
    {
        private const short createPartCommandId = 1;
        private const short createAssemblyCommandId = 2;
        private const short renameSelectedCommandId = 3;
        private const short exportStlCommandId = 4;

        public OriginFix() : base("Origin naming fix")
        {
        }

        protected override void Action(short command, short mode, object kompasObj)
        {
            KompasObject kompas = (KompasObject)kompasObj;
            bool isSuccess;
            switch (command)
            {
                case createPartCommandId:
                    isSuccess = DocHelpers.CreateNew(kompas, DocumentTypeEnum.ksDocumentPart);
                    break;
                case createAssemblyCommandId:
                    isSuccess = DocHelpers.CreateNew(kompas, DocumentTypeEnum.ksDocumentAssembly);
                    break;
                case renameSelectedCommandId:
                    string newName = kompas.ksReadString("New name", string.Empty);
                    isSuccess = DocHelpers.RenameSelectedPart(kompas, newName);
                    break;
                case exportStlCommandId:
                    isSuccess = ExportStl(kompas);
                    break;
                default:
                    isSuccess = false;
                    break;
            }
            if (!isSuccess)
            {
                kompas.ksMessage("Failed to create new part.");
            }
        }

        private static bool ExportStl(KompasObject kompas)
        {
            bool isSuccess = true;
            do
            {
                var doc = (ksDocument3D) kompas.ActiveDocument3D();
                if (doc == null)
                {
                    isSuccess = false;
                    break;
                }

                var formatParam = (ksAdditionFormatParam)doc.AdditionFormatParam();
                formatParam.Init();

                formatParam.format = (short)D3FormatConvType.format_STL;
                formatParam.formatBinary = false;
                formatParam.topolgyIncluded = false;
                formatParam.SetObjectsOptions((int)ksD3ConverterOptionsEnum.ksD3COBodyes, 1);
                formatParam.SetObjectsOptions((int)ksD3ConverterOptionsEnum.ksD3COSurfaces, 0);

                int stepType = 0;
                stepType |= (int)ksStepTypeEnum.ksSpaceStep;
                stepType |= (int)ksStepTypeEnum.ksDeviationStep;
                formatParam.stepType = stepType;

                formatParam.lengthUnits = (int)ksLengthUnitsEnum.ksLUnMM;
                formatParam.step = 0.1f;

                formatParam.angle = 3.0f  * Math.PI / 180;

                string fn = doc.fileName;

                doc.SaveAsToAdditionFormat(@"E:\RC\print_3d\cinelog_25\aaa.stl", formatParam);

            } while (false);
            return isSuccess;
        }

        // ReSharper disable once UnusedMember.Global
        // ReSharper disable once RedundantAssignment
        [return: MarshalAs(UnmanagedType.BStr)]
        public string ExternalMenuItem(short number, ref short itemType, ref short command)
        {
            string result = string.Empty;
            itemType = menuItemId;
            switch (number)
            {
                case 1:
                    result = "Create new part";
                    command = createPartCommandId;
                    break;
                case 2:
                    result = "Create new assembly";
                    command = createAssemblyCommandId;
                    break;
                case 3:
                    result = "Rename selected";
                    command = renameSelectedCommandId;
                    break;
                case 4:
                    result = "Export STL";
                    command = exportStlCommandId;
                    break;
                case 5:
                    command = -1;
                    itemType = menuEndId;
                    break;
            }
            return result;
        }
    }
}