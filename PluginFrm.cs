using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace Steps.NET
{
    public partial class PluginFrm : Form
    {
        public PluginFrm()
        {
            InitializeComponent();
            Properties = new Gear();
        }

        private static PluginFrm _instanse;

        public static PluginFrm Instanse
        {
            get
            {
                if (_instanse == null)
                {
                    _instanse = new PluginFrm();
                }
                return _instanse;
            }
        }

        public GearPlugin Plugin;

        private Gear Properties { get; set; }

        private static double ParseDouble(Control text)
        {
            double value = -1;
            try
            {
                value = Convert.ToDouble(text.Text);
            }
            catch (Exception ignore)
            {
                /* nothing */
            }
            return value;
        }

        private void diameterIn_TextChanged(object sender, EventArgs e)
        {
            Properties.DiameterIn = ParseDouble(diameterIn);
        }

        private void diameterOut_TextChanged(object sender, EventArgs e)
        {
            Properties.DiameterOut = ParseDouble(diameterOut);
        }

        private void shaftDiam_TextChanged(object sender, EventArgs e)
        {
            Properties.ShaftDiam = ParseDouble(shaftDiam);
        }

        private void thickness_TextChanged(object sender, EventArgs e)
        {
            Properties.Thickness = ParseDouble(thickness);
        }

        private void keywayDepth_TextChanged(object sender, EventArgs e)
        {
            Properties.KeywayDepth = ParseDouble(keywayDepth);
        }

        private void keywayWidth_TextChanged(object sender, EventArgs e)
        {
            Properties.KeywayWidth = ParseDouble(keywayWidth);
        }

        private void angle_TextChanged(object sender, EventArgs e)
        {
            Properties.Angle = ParseDouble(angle);
        }

        private void teethCount_TextChanged(object sender, EventArgs e)
        {
            short value = 0;
            try
            {
                value = Convert.ToByte(teethCount.Text);
            }
            catch (Exception ignore)
            {
            /* nothing */
            }
            Properties.TeethCount = value;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (Validate(Properties))
            {
                Plugin.CreateModel(Properties);
            }
        }

        private static bool Validate(Gear props)
        {
            return props.DiameterIn > 0 &&
                   props.DiameterOut > 0 &&
                   props.KeywayDepth >= 0 &&
                   props.KeywayWidth >= 0 &&
                   props.ShaftDiam > 0 &&
                   props.TeethCount > 1 &&
                   props.Angle >= 0 &&
                   props.Angle < 70 &&
                   props.Thickness > 0 &&
                   props.DiameterIn / 2 >
                   Math.Pow(Math.Pow(props.ShaftDiam / 2 + props.KeywayDepth, 2) + Math.Pow(props.KeywayWidth / 2, 2), 0.5) &&
                   props.DiameterIn < props.DiameterOut;
        }
    }

    public class Gear
    {
        private double _diameterIn;

        public double DiameterIn
        {
            get { return _diameterIn; }
            set { _diameterIn = value; }
        }

        private double _diameterOut;

        public double DiameterOut
        {
            get { return _diameterOut; }
            set { _diameterOut = value; }
        }

        private double _shaftDiam;

        public double ShaftDiam
        {
            get { return _shaftDiam; }
            set { _shaftDiam = value; }
        }

        private double _thickness;

        public double Thickness
        {
            get { return _thickness; }
            set { _thickness = value; }
        }
        private double _keywayDepth;

        public double KeywayDepth
        {
            get {  return _keywayDepth;}
            set { _keywayDepth = value; }
        }
        private double _keywayWidth;

        public double KeywayWidth
        {
            get { return _keywayWidth; }
            set { _keywayWidth = value; }
        }
        private double _angle;

        public double Angle
        {
            get { return _angle; }
            set { _angle = value; }
        }
        private short _teethCount;

        public short TeethCount
        {
            get { return _teethCount; }
            set { _teethCount = value; }
        }
    }
}
