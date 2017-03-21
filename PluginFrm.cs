using System;
using System.Windows.Forms;

namespace Steps.NET
{
    /// <summary>
    /// Форма задания параметров шестерни
    /// </summary>
    /// <seealso cref="System.Windows.Forms.Form" />
    public partial class PluginFrm : Form
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="PluginFrm"/> class.
        /// </summary>
        private PluginFrm()
        {
            InitializeComponent();
            Properties = new Gear();
        }

        /// <summary>
        /// Экземпляр класса
        /// </summary>
        private static PluginFrm _instanse;

        /// <summary>
        /// Получение экземпляра класса
        /// </summary>
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

        /// <summary>
        /// Ссылка на основной класс плагина
        /// </summary>
        public GearPlugin Plugin;

        /// <summary>
        /// Свойства шестерни
        /// </summary>
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

        private void button1_Click(object sender, EventArgs e)
        {
            Properties.DiameterIn = ParseDouble(diameterIn);
            Properties.DiameterOut = ParseDouble(diameterOut);
            Properties.ShaftDiam = ParseDouble(shaftDiam);
            Properties.Thickness = ParseDouble(thickness);
            Properties.KeywayDepth = ParseDouble(keywayDepth);
            Properties.KeywayWidth = ParseDouble(keywayWidth);
            Properties.Angle = ParseDouble(angle);
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
            if (Validate(Properties))
            {
                Plugin.CreateModel(Properties);
            }
        }

        private bool Validate(Gear props)
        {
            bool isValid = true;
            String errorMsg = "Следующие поля имеют некорректные значения:";
            if (props.DiameterIn <= 0)
            {
                errorMsg += "\nДиаметр окружности впадин: " + diameterIn.Text;
                isValid = false;
            }

            if (props.DiameterOut <= 0)
            {
                errorMsg += "\nДиаметр окружности вершин: " + diameterOut.Text;
                isValid = false;
            }

            if (props.KeywayDepth < 0)
            {
                errorMsg += "\nГлубина шпоночного паза: " + keywayDepth.Text;
                isValid = false;
            }

            if (props.KeywayWidth < 0)
            {
                errorMsg += "\nШирина шпоночного паза: " + keywayWidth.Text;
                isValid = false;
            }

            if (props.ShaftDiam <= 0)
            {
                errorMsg += "\nДиаметр отверстия: " + shaftDiam.Text;
                isValid = false;
            }

            if (props.TeethCount <= 2)
            {
                errorMsg += "\nКоличество зубьев: " + teethCount.Text;
                isValid = false;
            }

            if (props.Angle < 0)
            {
                errorMsg += "\nУгол наклона зубьев меньше 0: " + angle.Text;
                isValid = false;
            }

            if (props.Angle > 70)
            {
                errorMsg += "\nУгол наклона зубьев больше 70: " + angle.Text;
                isValid = false;
            }

            if (props.Thickness <= 0)
            {
                errorMsg += "\nТолщина детали: " + thickness.Text;
                isValid = false;
            }

            if (props.DiameterIn / 2 <=
                Math.Pow(Math.Pow(props.ShaftDiam / 2 + props.KeywayDepth, 2) + Math.Pow(props.KeywayWidth / 2, 2), 0.5))
            {
                errorMsg += "\nШпоночный паз выходит за пределы колеса";
                isValid = false;
            }

            if (props.DiameterIn >= props.DiameterOut)
            {
                errorMsg += "\nДиаметр впадин больше или равен диаметру вершин";
                isValid = false;
            }
            if (!isValid)
            {
                MessageBox.Show(errorMsg, "На форме есть ошибки");
            }
            return isValid;
        }
    }
}
