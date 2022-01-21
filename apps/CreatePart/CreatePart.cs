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
        public CreatePart() : base("New Part")
        {
        }

        protected override void Action(short command, short mode, object kompasObj)
        {
            KompasObject kompas = (KompasObject)kompasObj;
            if (!DocHelpers.CreateNew(kompas, DocHelpers.Doc3DType.Part))
            {
                kompas.ksMessage("Failed to create new part.");
            }
        }
    }
}