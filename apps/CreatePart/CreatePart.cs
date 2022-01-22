using System.Runtime.InteropServices;
using Kompas6API5;
using Base;
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
        private const short renamePartCommandId = 3;
        private const short renameAssemblyCommandId = 4;
        private const short renameSelectedCommandId = 5;

        public CreatePart() : base("Create New")
        {
        }

        protected override void Action(short command, short mode, object kompasObj)
        {
            KompasObject kompas = (KompasObject)kompasObj;
            bool isSuccess = true;
            switch (command)
            {
                case createPartCommandId:
                    isSuccess = DocHelpers.CreateNew(kompas, DocHelpers.Doc3DType.Part);
                    break;
                case createAssemblyCommandId:
                    isSuccess = DocHelpers.CreateNew(kompas, DocHelpers.Doc3DType.Assembly);
                    break;
                case renamePartCommandId:
                    DocHelpers.RenameTopPart(kompas, DocHelpers.Doc3DType.Part.ToString());
                    break;
                case renameAssemblyCommandId:
                    DocHelpers.RenameTopPart(kompas, DocHelpers.Doc3DType.Assembly.ToString());
                    break;
                case renameSelectedCommandId:
                    string newName = kompas.ksReadString("New name", string.Empty);
                    DocHelpers.RenameSelectedPart(kompas, newName);
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
                    result = "Part";
                    command = createPartCommandId;
                    break;
                case 2:
                    result = "Assembly";
                    command = createAssemblyCommandId;
                    break;
                case 3:
                    result = "Rename Part";
                    command = renamePartCommandId;
                    break;
                case 4:
                    result = "Rename Assembly";
                    command = renameAssemblyCommandId;
                    break;
                case 5:
                    result = "Rename Selected";
                    command = renameSelectedCommandId;
                    break;
                case 6:
                    command = -1;
                    itemType = menuEndId;
                    break;
            }
            return result;
        }
    }
}