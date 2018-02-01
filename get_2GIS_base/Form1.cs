using System;
using System.ComponentModel;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.IO;

namespace get_2GIS_base
{
    public partial class Form1 : Form
    {
        const int SW_SHOWMAXIMIZED = 3;
        private string fileToWrite = "";

        BackgroundWorker bgWorker = new BackgroundWorker();

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();

            saveFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            saveFileDialog1.FilterIndex = 1;
            saveFileDialog1.RestoreDirectory = true;

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                fileToWrite = saveFileDialog1.FileName;
            }
            else
            {
                MessageBox.Show("Файл для записи не выбран.", "User application", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            
            progressBar1.Value = 0;
            button1.Text = "Идет выгрузка...";

            bgWorker.WorkerReportsProgress = true;
            bgWorker.DoWork += new DoWorkEventHandler(RunGrym);
            bgWorker.ProgressChanged += new ProgressChangedEventHandler(bgWorker_ProgressChanged);
            bgWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bgWorker_RunWorkerCompleted);
            bgWorker.RunWorkerAsync();
            
        }
        
        private void RunGrym(object sender, DoWorkEventArgs evArgs)
        {         

            GrymCore.IGrym grymApp = null;
            GrymCore.IBaseReference baseRef = null;
            GrymCore.IBaseViewThread baseViewTread = null;
            GrymCore.ICommandLine cmdLine = null;
            GrymCore.IMapCoordinateTransformationGeo geoTransform = null;
            try
            {
                // Создаем объект приложения Grym.
                // Если приложение не было запущено, то при
                // первом же обращении к объекту оно запустится.
                grymApp = new GrymCore.Grym();
                // Получаем описание файла данных для заданного города
                // из коллекции описаний.
                baseRef = grymApp.BaseCollection.FindBase( cityFilter.Text.Trim() );
                if (baseRef == null)
                {
                    MessageBox.Show("Файл данных указанного города не найден.", "User application", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                // Получаем оболочку просмотра данных по описанию файла данных.
                // Если оболочка для запрошенной базы еще не открыта, то будет
                // запущена.
                baseViewTread = grymApp.GetBaseView(baseRef, true, false);
                if (baseViewTread == null)
                {
                    MessageBox.Show("Не удалось запустить оболочку просмотра данных.", "User application", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                StreamWriter sr = new StreamWriter(fileToWrite);
                sr.WriteLine("Город;Название_дома;Улица;Номер_дома;Координата_Х;Координата_Y;Назначение_Дома");

                GrymCore.IDatabase dataBase = baseViewTread.Database;
                GrymCore.ITable dataTable = dataBase.get_Table("grym_map_building");
                GrymCore.IDataRow dataRow = null;
                geoTransform = baseViewTread.Frame.Map.CoordinateTransformation as GrymCore.IMapCoordinateTransformationGeo;

                GrymCore.IMapPoint building_addr_position = null;
                int addr_count = 0;
                object building_city = null;
                object building_name = null, building_street = null, building_number = null, building_addr = null, building_purpose = null;

                int totalRecords = dataTable.RecordCount;
                for (int index = 1; index < totalRecords; index++)
                {
                    ((BackgroundWorker)sender).ReportProgress(100 * index / totalRecords);

                    dataRow = dataTable.GetRecord(index);                   
                    
                    building_city = dataRow.get_Value("city");
                    if ( building_city != null ) {

                        addr_count = int.Parse(dataRow.get_Value("addr_count").ToString());
                        if (addr_count > 0)
                        {
                            for (int ind = 1; ind <= addr_count; ind++)
                            {
                                building_name = dataRow.get_Value("name");
                                building_purpose = dataRow.get_Value("purpose");
                                building_street = dataRow.get_Value("street_" + ind);
                                building_number = dataRow.get_Value("number_" + ind);
                                building_addr = ((GrymCore.IDataRow)dataRow.get_Value("addr_" + ind)).get_Value("feature");

                                if (building_addr != null)
                                {
                                    building_addr_position = geoTransform.LocalToGeo((building_addr as GrymCore.IFeature).CenterPoint);

                                    sr.WriteLine(building_city + ";" + building_name + ";" + building_street + ";"
                                        + building_number + ";" + building_addr_position.X + ";" + building_addr_position.Y + ";" + building_purpose);
                                }
                            }
                        }

                    }
                }
                sr.Close();
            }
            catch (Exception e)
            {
                MessageBox.Show("Ошибка:" + e.ToString(), "User application", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            finally
            {
                if (grymApp != null)
                    Marshal.FinalReleaseComObject(grymApp);
                if (baseRef != null)
                    Marshal.FinalReleaseComObject(baseRef);
                if (baseViewTread != null)
                    Marshal.FinalReleaseComObject(baseViewTread);
                if (cmdLine != null)
                    Marshal.FinalReleaseComObject(cmdLine);

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        void bgWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage;
        }

        void bgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            progressBar1.Value = 100;
            button1.Text = "Выгрузить в файл";
        }
    }
}
