
using Kompas6API5;
using System;
using Microsoft.Win32;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Kompas6Constants3D;
using KAPITypes;
using reference = System.Int32;


namespace Steps.NET
{
    // ����� Plugin - ������� 3D
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class GearPlugin
    {
        private KompasObject kompas;
        private ksDocument3D doc3;
        private ksDocument2D doc;


        // ��� ����������
        [return: MarshalAs(UnmanagedType.BStr)]
        public string GetLibraryName()
        {
            return "������";
        }


        // �������� ������� ����������
        public void ExternalRunCommand([In] short command, [In] short mode,
            [In, MarshalAs(UnmanagedType.IDispatch)] object kompas_)
        {
            kompas = (KompasObject) kompas_;
            if (kompas != null)
            {
                doc3 = (ksDocument3D) kompas.ActiveDocument3D();
                if (doc3 == null || doc3.reference == 0)
                {
                    doc3 = (ksDocument3D) kompas.Document3D();
                    doc3.Create(true, true);

                    doc3.comment = "�������� ������";
                    doc3.drawMode = 3;
                    doc3.perspective = true;
                    doc3.UpdateDocumentParam();
                }
               switch (command)
                {
                    case 1:
                        PluginFrm.Instanse.Plugin = this;
                        PluginFrm.Instanse.ShowDialog();
                        break;
                }
            }
        }


        // ������������ ���� ����������
        [return: MarshalAs(UnmanagedType.BStr)]
        public string ExternalMenuItem(short number, ref short itemType, ref short command)
        {
            string result = string.Empty; //�� ��������� - ������ ������
            itemType = 1; //MENUITEM

            switch (number)
            {
                case 1:
                    result = "������� ��������";
                    command = 1;
                    break;
                case 2:
                    command = -1;
                    itemType = 3; // "ENDMENU"
                    break;
            }

            return result;
        }

        private static double DegToRad(double angle)
        {
            return Math.PI * angle / 180.0;
        }

        //�������� ������
        public void createModel(Gear properties)
        {
            int z = properties.TeethCount;
            double beta = properties.Angle;
            double Lm = properties.Thickness;
            // ������� ��������� ��� ���
            double Dv = properties.ShaftDiam;
            //Dv = 150;
            // ������ �������� � ������ ������ ��������� �������
            double b_k = Lm;
            // ������� ��������
            // ������� �����, ������������ �������� � ������
            // ������� �����
            double d_ak = properties.DiameterOut; // ������� �������� 
            double d_fk = properties.DiameterIn; // ������� ������
            double d_k = (d_ak + d_fk)/2; // ����������� ������� ������
            double alfa1 = 0.0;
            double alfa2 = 0.0;
            // ���������� ������������� ����������
            ksPart iPart = doc3.GetPart((int) Part_Type.pNew_Part) as ksPart; //����� ��������
            if (iPart != null)
            {
                // ���������� ������������� ����������
                ksEntity PlaneXOY = iPart.GetDefaultEntity((short) Obj3dType.o3d_planeXOY) as ksEntity;
                ksEntity PlaneXOZ = iPart.GetDefaultEntity((short) Obj3dType.o3d_planeXOZ) as ksEntity;
                ksEntity PlaneYOZ = iPart.GetDefaultEntity((short) Obj3dType.o3d_planeYOZ) as ksEntity;
                // ��������� ������ (�������� ������� ������� ������)
                ksEntity iSketchEntity = iPart.NewEntity((short) Obj3dType.o3d_sketch) as ksEntity;
                if (iSketchEntity != null)
                {
                    // ��������� ���������� ������
                    ksSketchDefinition iSketchDef = iSketchEntity.GetDefinition() as ksSketchDefinition;
                    if (iSketchDef != null)
                    {
                        if (PlaneXOY != null)
                        {
                            // ������������� ���������,
                            // �� ������� ��������� �����
                            iSketchDef.SetPlane(PlaneXOY);

                            iSketchEntity.Create();

                            // ��������� ������� �������������� ������
                            // doc � ��������� �� ��������� ksDocument2D
                            doc = iSketchDef.BeginEdit() as ksDocument2D;
                            if (doc != null)
                            {
                                doc.ksLineSeg(-Lm / 2, 00, Lm / 2, 0, 3); //3 - ������ �����
                                doc.ksLineSeg(Lm / 2, Dv / 2, -Lm / 2, Dv / 2, 1);
                                doc.ksLineSeg(Lm / 2, Dv / 2, Lm / 2, d_ak / 2, 1);
                                doc.ksLineSeg(-Lm / 2, d_ak / 2, -Lm / 2, Dv / 2, 1);
                                doc.ksLineSeg(-Lm/2, d_ak / 2, Lm/2, d_ak / 2, 1);
                            }
                            iSketchDef.EndEdit();
                        }
                    }
                }
                // ��������� ������� �������� ��������
                ksEntity iBaseRotatedEntity = iPart.NewEntity((short) Obj3dType.o3d_baseRotated) as ksEntity;
                // ��������� ���������� ����� � ���������� �������
                ksColorParam Color = iBaseRotatedEntity.ColorParam() as ksColorParam;
                Color.specularity = 0.8;
                Color.shininess = 1;
                if (iBaseRotatedEntity != null)
                {
                    // ��������� ���������� ��������
                    ksBaseRotatedDefinition iBaseRotatedDef =
                        iBaseRotatedEntity.GetDefinition() as ksBaseRotatedDefinition;
                    if (iBaseRotatedDef != null)
                    {
                        // ��������� ���������� ��������
                        iBaseRotatedDef.SetThinParam(false, (short) Direction_Type.dtNormal, 1, 1);
                        iBaseRotatedDef.SetSideParam(true, 360.0);
                        iBaseRotatedDef.toroidShapeType = false;
                        iBaseRotatedDef.SetSketch(iSketchEntity);
                        // ������� �������� ��������
                        // ��������� � ��������� ��������� ������
                        iBaseRotatedEntity.Create();
                    }
                }
                ksEntity iSketch1Entity = iPart.NewEntity((short) Obj3dType.o3d_sketch) as ksEntity;
                if (iSketch1Entity != null)
                {
                    ksSketchDefinition iSketch1Def = iSketch1Entity.GetDefinition() as ksSketchDefinition;
                    if (iSketch1Def != null)
                    {
                        if (PlaneYOZ != null)
                        {
                            // ��������� ����� �� ��������� XOY
                            iSketch1Def.SetPlane(PlaneXOY);
                            iSketch1Entity.Create();
                            doc = iSketch1Def.BeginEdit() as ksDocument2D;
                            if (doc != null)
                            {
                                // ����������� � ������ � 4 ����������
                                double width = properties.KeywayWidth;
                                doc.ksLineSeg(Lm / 2, -width / 2, -Lm / 2, -width / 2, 1);
                                doc.ksLineSeg(-Lm / 2, width / 2, -Lm / 2, -width / 2, 1);
                                doc.ksLineSeg(-Lm / 2, width / 2, Lm / 2, width / 2, 1);
                                doc.ksLineSeg(Lm / 2, width / 2, Lm / 2, -width / 2, 1);
                            }
                            iSketch1Def.EndEdit();
                        }
                    }
                }
                // ��������� �������� �������� �������������
                ksEntity iCutExtrusion = iPart.NewEntity((short) Obj3dType.o3d_cutExtrusion) as ksEntity;
                if (iCutExtrusion != null)
                {
                    // ��������� ���������� ���������
                    ksCutExtrusionDefinition iCutExtrusionDef =
                        iCutExtrusion.GetDefinition() as ksCutExtrusionDefinition;
                    if (iCutExtrusionDef != null)
                    {
                        double depth = properties.KeywayDepth;
                        // ��������� ���������� 
                        iCutExtrusionDef.SetSketch(iSketch1Entity);
                        // �����������
                        iCutExtrusionDef.directionType = (short) Direction_Type.dtNormal;
                        // �������� ��������� �� ������� �� �����������
                        iCutExtrusionDef.SetSideParam(true, (short)ksEndTypeEnum.etBlind, depth + Dv / 2, 0, false);
                        iCutExtrusionDef.SetSideParam(false, (short)ksEndTypeEnum.etBlind, depth + Dv / 2, 0, false);
                        iCutExtrusionDef.SetThinParam(false, 0, 0, 0);
                        // ������� ��������� � �����
                        iCutExtrusion.Create();
                    }
                }
                // ��������� ��������� ���������
                ksEntity iOffsetPlaneEntity = iPart.NewEntity((short) Obj3dType.o3d_planeOffset) as ksEntity;
                if (iOffsetPlaneEntity != null)
                {
                    // ��������� ���������� ��������� ���������
                    ksPlaneOffsetDefinition iOffsetPlaneDef =
                        iOffsetPlaneEntity.GetDefinition() as ksPlaneOffsetDefinition;
                    if (iOffsetPlaneDef != null)
                    {
                        // ��������, ������� ��������� � ������ ��������� ��������
                        iOffsetPlaneDef.offset = b_k / 2;
                        iOffsetPlaneDef.SetPlane(PlaneYOZ);
                        iOffsetPlaneDef.direction = false;
                        // ������ ��������� �������
                        iOffsetPlaneEntity.hidden = true;
                        // ������� ��������������� ���������
                        iOffsetPlaneEntity.Create();
                    }
                }
                // ����� ������� ������ ����� �������
                ksEntity iSketch2Entity = iPart.NewEntity((short) Obj3dType.o3d_sketch) as ksEntity;
                if (iSketch2Entity != null)
                {
                    ksSketchDefinition iSketch2Def = iSketch2Entity.GetDefinition() as ksSketchDefinition;
                    if (iSketch2Def != null)
                    {
                        // ������� ��������� � ��������������� iOffsetPlaneEntity
                        iSketch2Def.SetPlane(iOffsetPlaneEntity);
                        iSketch2Entity.Create();
                        doc = iSketch2Def.BeginEdit() as ksDocument2D;
                        alfa1 = 360.0 / z;
                        doc.ksMtr(0, 0, 90, 1, 1);
                        // ������������ ����������� ������
                        // ������ ��������� ��� ��������
                        // ����� ������� ���� �� ���� ������
                        doc.ksArcBy3Points(-(d_ak / 2 + 0.1) * Math.Sin(DegToRad(0)),
                            -(d_ak / 2 + 0.1) * Math.Cos(DegToRad(0)),
                            -d_k / 2 * Math.Sin(DegToRad(alfa1 / 8)), -d_k / 2 * Math.Cos(DegToRad(alfa1 / 8)),
                            -d_fk / 2 * Math.Sin(DegToRad(alfa1 / 4)), -d_fk / 2 * Math.Cos(DegToRad(alfa1 / 4)), 1);
                        doc.ksArcByPoint(0, 0, d_fk / 2, -d_fk / 2 * Math.Sin(DegToRad(alfa1 / 4)),
                            -d_fk / 2 * Math.Cos(DegToRad(alfa1 / 4)),
                            -d_fk / 2 * Math.Sin(DegToRad(alfa1 / 2)), -d_fk / 2 * Math.Cos(DegToRad(alfa1 / 2)), -1, 1);
                        doc.ksArcBy3Points(-d_fk / 2 * Math.Sin(DegToRad(alfa1 / 2)),
                            -d_fk / 2 * Math.Cos(DegToRad(alfa1 / 2)),
                            -d_k / 2 * Math.Sin(DegToRad(0.625 * alfa1)), -d_k / 2 * Math.Cos(DegToRad(0.625 * alfa1)),
                            -(d_ak / 2 + 0.1) * Math.Sin(DegToRad(0.75 * alfa1)),
                            -(d_ak / 2 + 0.1) * Math.Cos(DegToRad(0.75 * alfa1)), 1);
                        doc.ksArcBy3Points(-(d_ak / 2 + 0.1) * Math.Sin(DegToRad(0.75 * alfa1)),
                            -(d_ak / 2 + 0.1) * Math.Cos(DegToRad(0.75 * alfa1)),
                            -(d_ak / 2 + 2) * Math.Sin(DegToRad(0.375 * alfa1)),
                            -(d_ak / 2 + 2) * Math.Cos(DegToRad(0.375 * alfa1)),
                            -(d_ak / 2 + 0.1) * Math.Sin(DegToRad(0)), -(d_ak / 2 + 0.1) * Math.Cos(DegToRad(0)), 1);
                        doc.ksDeleteMtr();
                        iSketch2Def.EndEdit();
                    }
                }
                // ��������� ������� ������ ������ ����� �������
                ksEntity iSketch3Entity = iPart.NewEntity((short) Obj3dType.o3d_sketch) as ksEntity;
                if (iSketch3Entity != null)
                {
                    ksSketchDefinition iSketch3Def = iSketch3Entity.GetDefinition() as ksSketchDefinition;
                    if (iSketch3Def != null)
                    {
                        // ������ �� ��������� YOZ
                        iSketch3Def.SetPlane(PlaneYOZ);
                        iSketch3Entity.Create();
                        doc = iSketch3Def.BeginEdit() as ksDocument2D;
                        alfa2 = -(180 / Math.PI * (b_k * Math.Tan(Math.PI * (beta) / 180) / d_k));
                        doc.ksMtr(0, 0, 90, 1, 1);
                        // ������������ ����������� ������
                        // ������ ��������� ��� ��������
                        // ����� ������� ���� �� ���� ������
                        doc.ksArcBy3Points(-(d_ak / 2 + 0.1) * Math.Sin(DegToRad(alfa2)),
                            -(d_ak / 2 + 0.1) * Math.Cos(DegToRad(alfa2)),
                            -d_k / 2 * Math.Sin(DegToRad(alfa1 / 8 + alfa2)),
                            -d_k / 2 * Math.Cos(DegToRad(alfa1 / 8 + alfa2)),
                            -d_fk / 2 * Math.Sin(DegToRad(alfa1 / 4 + alfa2)),
                            -d_fk / 2 * Math.Cos(DegToRad(alfa1 / 4 + alfa2)), 1);
                        doc.ksArcByPoint(0, 0, d_fk / 2, -d_fk / 2 * Math.Sin(DegToRad(alfa1 / 4 + alfa2)),
                            -d_fk / 2 * Math.Cos(DegToRad(alfa1 / 4 + alfa2)),
                            -d_fk / 2 * Math.Sin(DegToRad(alfa1 / 2 + alfa2)),
                            -d_fk / 2 * Math.Cos(DegToRad(alfa1 / 2 + alfa2)), -1, 1);
                        doc.ksArcBy3Points(-d_fk / 2 * Math.Sin(DegToRad(alfa1 / 2 + alfa2)),
                            -d_fk / 2 * Math.Cos(DegToRad(alfa1 / 2 + alfa2)),
                            -d_k / 2 * Math.Sin(DegToRad(0.625 * alfa1 + alfa2)),
                            -d_k / 2 * Math.Cos(DegToRad(0.625 * alfa1 + alfa2)),
                            -(d_ak / 2 + 0.1) * Math.Sin(DegToRad(0.75 * alfa1 + alfa2)),
                            -(d_ak / 2 + 0.1) * Math.Cos(DegToRad(0.75 * alfa1 + alfa2)), 1);
                        doc.ksArcBy3Points(-(d_ak / 2 + 0.1) * Math.Sin(DegToRad(0.75 * alfa1 + alfa2)),
                            -(d_ak / 2 + 0.1) * Math.Cos(DegToRad(0.75 * alfa1 + alfa2)),
                            -(d_ak / 2 + 2) * Math.Sin(DegToRad(0.375 * alfa1 + alfa2)),
                            -(d_ak / 2 + 2) * Math.Cos(DegToRad(0.375 * alfa1 + alfa2)),
                            -(d_ak / 2 + 0.1) * Math.Sin(DegToRad(alfa2)), -(d_ak / 2 + 0.1) * Math.Cos(DegToRad(alfa2)),
                            1);
                        doc.ksDeleteMtr();
                        iSketch3Def.EndEdit();
                    }
                }
                // ������ ��������� ���������
                ksEntity iOffsetPlane1Entity = iPart.NewEntity((short) Obj3dType.o3d_planeOffset) as ksEntity;
                if (iOffsetPlane1Entity != null)
                {
                    ksPlaneOffsetDefinition iOffsetPlane1Def =
                        iOffsetPlane1Entity.GetDefinition() as ksPlaneOffsetDefinition;
                    if (iOffsetPlane1Def != null)
                    {
                        // �������� �������� �� ��
                        iOffsetPlane1Def.offset = b_k / 2;
                        // ����������� ���������������
                        iOffsetPlane1Def.direction = true;
                        iOffsetPlane1Def.SetPlane(PlaneYOZ);
                        // ������ ��������� �������
                        iOffsetPlane1Entity.hidden = true;
                        // ������� ��������� ���������
                        iOffsetPlane1Entity.Create();
                    }
                }
                // ������ (���������) ����� ������ ����� �������
                ksEntity iSketch4Entity = iPart.NewEntity((short) Obj3dType.o3d_sketch) as ksEntity;
                if (iSketch4Entity != null)
                {
                    ksSketchDefinition iSketch4Def = iSketch4Entity.GetDefinition() as ksSketchDefinition;
                    if (iSketch4Def != null)
                    {
                        // ������� ��������� � ������ ��� ��������� ���������
                        iSketch4Def.SetPlane(iOffsetPlane1Entity);
                        iSketch4Entity.Create();
                        doc = iSketch4Def.BeginEdit() as ksDocument2D;
                        alfa2 = -(180 / Math.PI * (2 * b_k * Math.Tan(Math.PI * (beta) / 180) / d_k));
                        doc.ksMtr(0, 0, 90, 1, 1);
                        // ������������ ����������� ������
                        // ������ ��������� ��� ��������
                        // ����� ������� ���� �� ���� ������
                        doc.ksArcBy3Points(-(d_ak / 2 + 0.1) * Math.Sin(DegToRad(alfa2)),
                            -(d_ak / 2 + 0.1) * Math.Cos(DegToRad(alfa2)),
                            -d_k / 2 * Math.Sin(DegToRad(alfa1 / 8 + alfa2)),
                            -d_k / 2 * Math.Cos(DegToRad(alfa1 / 8 + alfa2)),
                            -d_fk / 2 * Math.Sin(DegToRad(alfa1 / 4 + alfa2)),
                            -d_fk / 2 * Math.Cos(DegToRad(alfa1 / 4 + alfa2)), 1);
                        doc.ksArcByPoint(0, 0, d_fk / 2, -d_fk / 2 * Math.Sin(DegToRad(alfa1 / 4 + alfa2)),
                            -d_fk / 2 * Math.Cos(DegToRad(alfa1 / 4 + alfa2)),
                            -d_fk / 2 * Math.Sin(DegToRad(alfa1 / 2 + alfa2)),
                            -d_fk / 2 * Math.Cos(DegToRad(alfa1 / 2 + alfa2)), -1, 1);
                        doc.ksArcBy3Points(-d_fk / 2 * Math.Sin(DegToRad(alfa1 / 2 + alfa2)),
                            -d_fk / 2 * Math.Cos(DegToRad(alfa1 / 2 + alfa2)),
                            -d_k / 2 * Math.Sin(DegToRad(0.625 * alfa1 + alfa2)),
                            -d_k / 2 * Math.Cos(DegToRad(0.625 * alfa1 + alfa2)),
                            -(d_ak / 2 + 0.1) * Math.Sin(DegToRad(0.75 * alfa1 + alfa2)),
                            -(d_ak / 2 + 0.1) * Math.Cos(DegToRad(0.75 * alfa1 + alfa2)), 1);
                        doc.ksArcBy3Points(-(d_ak / 2 + 0.1) * Math.Sin(DegToRad(0.75 * alfa1 + alfa2)),
                            -(d_ak / 2 + 0.1) * Math.Cos(DegToRad(0.75 * alfa1 + alfa2)),
                            -(d_ak / 2 + 2) * Math.Sin(DegToRad(0.375 * alfa1 + alfa2)),
                            -(d_ak / 2 + 2) * Math.Cos(DegToRad(0.375 * alfa1 + alfa2)),
                            -(d_ak / 2 + 0.1) * Math.Sin(DegToRad(alfa2)), -(d_ak / 2 + 0.1) * Math.Cos(DegToRad(alfa2)),
                            1);
                        doc.ksDeleteMtr();
                        iSketch4Def.EndEdit();
                    }
                }
                // ��������� �������� �������� �� ��������
                ksEntity iCutLoftEntity = iPart.NewEntity((short) Obj3dType.o3d_cutLoft) as ksEntity;
                if (iCutLoftEntity != null)
                {
                    // ��������� ���������� �������� �� ��������
                    ksCutLoftDefinition iCutLoftDef = iCutLoftEntity.GetDefinition() as ksCutLoftDefinition;
                    if (iCutLoftDef != null)
                    {
                        // ��������� ������� ksEntityCollection
                        // ��������� ������� ��� ��������� �� ��������
                        ksEntityCollection Collect = iCutLoftDef.Sketchs() as ksEntityCollection;
                        // ��������� ������ � ��������
                        Collect.Add(iSketch2Entity);
                        Collect.Add(iSketch3Entity);
                        Collect.Add(iSketch4Entity);
                        // ������� �������� �� ��������
                        // ��������� � ������ ����� ����� ������� � ����� ������
                        iCutLoftEntity.Create();
                    }
                }
                // ��������� ��������������� ��� �� ����������� ���� ����������
                ksEntity iAxis = iPart.NewEntity((short) Obj3dType.o3d_axis2Planes) as ksEntity;
                if (iAxis != null)
                {
                    // ��������� ���������� ��������������� ���
                    // �� ����������� ����������
                    ksAxis2PlanesDefinition iAxis2PlDef = iAxis.GetDefinition() as ksAxis2PlanesDefinition;
                    if (iAxis2PlDef != null)
                    {
                        // ������ ���������
                        iAxis2PlDef.SetPlane(1, PlaneXOZ);
                        iAxis2PlDef.SetPlane(2, PlaneXOY);
                        // ������ ��� ���������
                        iAxis.hidden = true;
                        // ������� ��������������� ���
                        iAxis.Create();
                    }
                }
                // ��������� �������� ������ �� ��������������� �����
                ksEntity iCircularCopy = iPart.NewEntity((short) Obj3dType.o3d_circularCopy) as ksEntity;
                if (iCircularCopy != null)
                {
                    // ��������� ���������� �������� ����������� �� �������
                    ksCircularCopyDefinition iCirCopyDef = iCircularCopy.GetDefinition() as ksCircularCopyDefinition;
                    if (iCirCopyDef != null)
                    {
                        // ��������� �������� ��� �����������
                        ksEntityCollection Collect1 = iCirCopyDef.GetOperationArray() as ksEntityCollection;
                        // �������� ����� ���� ���� � ��������� ����
                        Collect1.Add(iCutLoftEntity);
                        // ���������� �����, ����� ���������� ������
                        iCirCopyDef.count2 = z;
                        iCirCopyDef.factor2 = true;
                        // ��� �����������
                        iCirCopyDef.SetAxis(iAxis);
                        // ������� ��������������� ������ � ������ ������!
                        iCircularCopy.Create();
                    }
                }
            }
        }

        #region COM Registration

        // ��� ������� ����������� ��� ����������� ������ ��� COM
        // ��� ��������� � ����� ������� ���������� ������ Kompas_Library,
        // ������� ������������� � ���, ��� ����� �������� ����������� ������,
        // � ����� �������� ��� InprocServer32 �� ������, � ��������� ����.
        // ��� ��� �������� ��� ����, ����� ����� ����������� ����������
        // ���������� �� ������� ActiveX.
        [ComRegisterFunction]
        public static void RegisterKompasLib(Type t)
        {
            try
            {
                RegistryKey regKey = Registry.LocalMachine;
                string keyName = @"SOFTWARE\Classes\CLSID\{" + t.GUID.ToString() + "}";
                regKey = regKey.OpenSubKey(keyName, true);
                regKey.CreateSubKey("Kompas_Library");
                regKey = regKey.OpenSubKey("InprocServer32", true);
                regKey.SetValue(null,
                    System.Environment.GetFolderPath(Environment.SpecialFolder.System) + @"\mscoree.dll");
                regKey.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("��� ����������� ������ ��� COM-Interop ��������� ������:\n{0}", ex));
            }
        }

        // ��� ������� ������� ������ Kompas_Library �� �������
        [ComUnregisterFunction]
        public static void UnregisterKompasLib(Type t)
        {
            RegistryKey regKey = Registry.LocalMachine;
            string keyName = @"SOFTWARE\Classes\CLSID\{" + t.GUID.ToString() + "}";
            RegistryKey subKey = regKey.OpenSubKey(keyName, true);
            subKey.DeleteSubKey("Kompas_Library");
            subKey.Close();
        }

        #endregion
    }
}
