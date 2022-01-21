using Kompas6API5;
using Base;
using Kompas6Constants3D;

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
            bool isSuccess;
            do
            {
                // ReSharper disable once NotAccessedVariable
                Document3D d3d;
                isSuccess = DocHelpers.CreateDoc3D(out d3d, kompas, DocHelpers.Doc3DType.Part);
                var topPart = (ksPart)d3d.GetPart((int)Part_Type.pTop_Part);
                string name = topPart.name;
                kompas.ksMessage(name);
                topPart.name = "Part";
                topPart.Update();

            } while (false);

            if (isSuccess)
            {
                kompas.ksMessage("New part created.");
            }
            else
            {
                kompas.ksMessage("Failed to create new part.");
            }
        }
    }
}