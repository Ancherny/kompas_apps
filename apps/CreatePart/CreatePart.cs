using Kompas6API5;
using Kompas6Constants3D;
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
            bool isSuccess;
            do
            {
                Document3D d3d;
                isSuccess = DocHelpers.CreateDoc3D(out d3d, kompas, DocHelpers.Doc3DType.Part);
                var topPart = (ksPart)d3d.GetPart((int)Part_Type.pTop_Part);
                topPart.name = "Part";
                topPart.Update();

                DocHelpers.SetPartFirstEntityName(topPart, "Origin", Obj3dType.o3d_pointCS);
                DocHelpers.SetPartFirstEntityName(topPart, "Plane_XY", Obj3dType.o3d_planeXOY);
                DocHelpers.SetPartFirstEntityName(topPart, "Plane_XZ", Obj3dType.o3d_planeXOZ);
                DocHelpers.SetPartFirstEntityName(topPart, "Plane_YZ", Obj3dType.o3d_planeYOZ);
                DocHelpers.SetPartFirstEntityName(topPart, "Axis_X", Obj3dType.o3d_axisOX);
                DocHelpers.SetPartFirstEntityName(topPart, "Axis_X", Obj3dType.o3d_axisOY);
                DocHelpers.SetPartFirstEntityName(topPart, "Axis_X", Obj3dType.o3d_axisOZ);

            } while (false);

            if (!isSuccess)
            {
                kompas.ksMessage("Failed to create new part.");
            }
        }
    }
}