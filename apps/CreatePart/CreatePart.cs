using System.Runtime.InteropServices;
using Kompas6API5;
using Base;
using Kompas6Constants;

namespace RunCommands
{
    /// <summary>
    /// Create new part with english namings
    /// </summary>
    ///
    // ReSharper disable once UnusedType.Global
    public class CreatePart : BaseCommand
    {
        private const short createPartCommandId = 1;
        private const short createAssemblyCommandId = 2;
        private const short renameSelectedCommandId = 3;

        public CreatePart() : base("Origin naming fix")
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
                default:
                    isSuccess = false;
                    break;
            }
            if (!isSuccess)
            {
                kompas.ksMessage("Failed to create new part.");
            }
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
                    result = "Rename Selected";
                    command = renameSelectedCommandId;
                    break;
                case 4:
                    command = -1;
                    itemType = menuEndId;
                    break;
            }
            return result;
        }
    }
}