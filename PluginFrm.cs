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

        private static PluginFrm instanse;

        public static PluginFrm Instanse
        {
            get
            {
                if (instanse == null)
                {
                    instanse = new PluginFrm();
                }
                return instanse;
            }
        }

        public GearPlugin Plugin;

        private Gear Properties { get; set; }

        private double ParseDouble(TextBox text)
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
            short Value = 0;
            try
            {
                Value = Convert.ToByte(teethCount.Text);
            }
            catch (Exception ignore)
            {
            /* nothing */
            }
            Properties.TeethCount = Value;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (Validate(Properties))
            {
                Plugin.createModel(Properties);
            }
        }

        private bool Validate(Gear props)
        {
            return props.DiameterIn > 0 &&
                   props.DiameterOut > 0 &&
                   props.KeywayDepth > 0 &&
                   props.KeywayWidth > 0 &&
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
        public double DiameterIn;
        public double DiameterOut;
        public double ShaftDiam;
        public double Thickness;
        public double KeywayDepth;
        public double KeywayWidth;
        public double Angle;
        public short TeethCount;
    }
}
