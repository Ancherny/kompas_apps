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

    public static bool CreateNew(KompasObject kompas, Doc3DType docType)
    {
        bool isSuccess;
        do
        {
            isSuccess = CreateDoc3D(out _, kompas, docType);
            if (!isSuccess)
            {
                break;
            }

            ksPart topPart = RenameTopPart(kompas, docType.ToString());
            RenamePartOrigin(topPart);

        } while (false);

        return isSuccess;
    }

    public static ksPart RenameTopPart(KompasObject kompas, [NotNull] string name)
    {
        Document3D d3d = (Document3D)kompas.ActiveDocument3D();
        var topPart = (ksPart)d3d.GetPart((int)Part_Type.pTop_Part);
        topPart.name = name;
        topPart.Update();
        RenamePartOrigin(topPart);
        return topPart;
    }

    public static bool RenameSelectedPart(KompasObject kompas, [NotNull] string name)
    {
        bool isSuccess = true;
        do
        {
            Document3D d3d = (Document3D)kompas.ActiveDocument3D();
            var selMgr = (SelectionMng)d3d.GetSelectionMng();
            var selPart = (ksPart)selMgr.First();
            if (selPart == null)
            {
                isSuccess = false;
                break;
            }
            selPart.name = name;
            selPart.Update();
            RenamePartOrigin(selPart);

        } while (false);

        return isSuccess;
    }

    private static void SetPartFirstEntityName([NotNull] ksPart part, [NotNull] string name, Obj3dType entityType)
    {
        var collection = (ksEntityCollection)part.EntityCollection((short)entityType);
        var entity = (ksEntity)collection.GetByIndex(0);
        entity.name = name;
        entity.Update();
    }

    // ReSharper disable once MemberCanBePrivate.Global
    public static void RenamePartOrigin(ksPart part)
    {
        SetPartFirstEntityName(part, "Origin", Obj3dType.o3d_pointCS);
        SetPartFirstEntityName(part, "Plane_XY", Obj3dType.o3d_planeXOY);
        SetPartFirstEntityName(part, "Plane_XZ", Obj3dType.o3d_planeXOZ);
        SetPartFirstEntityName(part, "Plane_YZ", Obj3dType.o3d_planeYOZ);
        SetPartFirstEntityName(part, "Axis_X", Obj3dType.o3d_axisOX);
        SetPartFirstEntityName(part, "Axis_Y", Obj3dType.o3d_axisOY);
        SetPartFirstEntityName(part, "Axis_Z", Obj3dType.o3d_axisOZ);
    }
}