using JetBrains.Annotations;
using Kompas6API5;
using KAPITypes;
using Kompas6Constants;
using Kompas6Constants3D;

namespace Base;

public static class DocHelpers
{
    public enum Doc3DType
    {
        Part,
        Assembly
    }

    // ReSharper disable once UnusedMember.Global
    public static void CreateDoc2D(object kompasObj)
    {
        var kompas = (KompasObject)kompasObj;

        var dp = (DocumentParam)kompas.GetParamStruct((short)StructType2DEnum.ko_DocumentParam);
        dp.Init();
        dp.type = (short)DocType.lt_DocPart3D;

        var d2d = (Document2D)kompas.Document2D();
        d2d.ksCreateDocument(dp);
    }

    // ReSharper disable once MemberCanBePrivate.Global
    public static bool CreateDoc3D(out Document3D d3d, [NotNull] KompasObject kompas, Doc3DType doc3DType)
    {
        d3d = null;
        bool isSuccess;
        d3d = (Document3D)kompas.Document3D();
        switch (doc3DType)
        {
            case Doc3DType.Part:
                isSuccess = d3d.Create();
                break;
            case Doc3DType.Assembly:
                isSuccess = d3d.Create(false, false);
                break;
            default:
                isSuccess = false;
                break;
        }
        return isSuccess;
    }

    private static void SetPartFirstEntityName([NotNull] ksPart part, [NotNull] string name, Obj3dType entityType)
    {
        var collection = (ksEntityCollection)part.EntityCollection((short)entityType);
        var entity = (ksEntity)collection.GetByIndex(0);
        entity.name = name;
        entity.Update();
    }

    public static bool CreateNew(KompasObject kompas, Doc3DType docType)
    {
        bool isSuccess;
        do
        {
            Document3D d3d;
            isSuccess = CreateDoc3D(out d3d, kompas, docType);
            var topPart = (ksPart)d3d.GetPart((int)Part_Type.pTop_Part);
            topPart.name = docType.ToString();
            topPart.Update();

            SetPartFirstEntityName(topPart, "Origin", Obj3dType.o3d_pointCS);
            SetPartFirstEntityName(topPart, "Plane_XY", Obj3dType.o3d_planeXOY);
            SetPartFirstEntityName(topPart, "Plane_XZ", Obj3dType.o3d_planeXOZ);
            SetPartFirstEntityName(topPart, "Plane_YZ", Obj3dType.o3d_planeYOZ);
            SetPartFirstEntityName(topPart, "Axis_X", Obj3dType.o3d_axisOX);
            SetPartFirstEntityName(topPart, "Axis_X", Obj3dType.o3d_axisOY);
            SetPartFirstEntityName(topPart, "Axis_X", Obj3dType.o3d_axisOZ);

        } while (false);

        return isSuccess;
    }
}