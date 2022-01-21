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

        public CreatePart() : base("Create New")
        {
        }

        protected override void Action(short command, short mode, object kompasObj)
        {
            KompasObject kompas = (KompasObject)kompasObj;
            bool isSuccess;
            switch (command)
            {
                case createPartCommandId:
                    isSuccess = DocHelpers.CreateNew(kompas, DocHelpers.Doc3DType.Part);
                    break;
                case createAssemblyCommandId:
                    isSuccess = DocHelpers.CreateNew(kompas, DocHelpers.Doc3DType.Assembly);
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
                    command = -1;
                    itemType = menuEndId;
                    break;
            }
            return result;
        }
    }
}