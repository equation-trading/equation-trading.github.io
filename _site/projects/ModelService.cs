using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using System.Runtime.InteropServices;


namespace ModelService
{
    public class ModelService
    {
        [DllImport("CopulaLibrary.dll", CallingConvention = CallingConvention.Cdecl, EntryPoint = "CopulaService_ArchimedeanCopulaCalibration")]
        extern static void CopulaService_ArchimedeanCopulaCalibration(ref int type, ref double theta, ref int dim, ref int sample_size, double[] output);
        [DllImport("CopulaLibrary.dll", CallingConvention = CallingConvention.Cdecl, EntryPoint = "CopulaService_GaussianCopulaCalibration")]
        extern static void CopulaService_GaussianCopulaCalibration(double[] correlation, ref int dim, ref int sample_size, double[] output);
        [DllImport("CopulaLibrary.dll", CallingConvention = CallingConvention.Cdecl, EntryPoint = "CopulaService_StudentCopulaCalibration")]
        extern static void CopulaService_StudentCopulaCalibration(ref double degreeOfFreedom, double[] correlation, ref int dim, ref int sample_size, double[] output);
        [DllImport("CopulaLibrary.dll", CallingConvention = CallingConvention.Cdecl, EntryPoint = "CopulaService_ArchimedeanCopulaSimulation")]
        extern static void CopulaService_ArchimedeanCopulaSimulation(ref int type, ref double theta, ref int dim, ref int sample_size, double[] output);
        [DllImport("CopulaLibrary.dll", CallingConvention = CallingConvention.Cdecl, EntryPoint = "CopulaService_GaussianCopulaSimulation")]
        extern static void CopulaService_GaussianCopulaSimulation(double[] correlation, ref int dim, ref int sample_size, double[] output);
        [DllImport("CopulaLibrary.dll", CallingConvention = CallingConvention.Cdecl, EntryPoint = "CopulaService_StudentCopulaSimulation")]
        extern static void CopulaService_StudentCopulaSimulation(ref double degreeOfFreedom, double[] correlation, ref int dim, ref int sample_size, double[] output);
        [DllImport("CopulaLibrary.dll", CallingConvention = CallingConvention.Cdecl, EntryPoint = "SmoothingService_KernelSmoothing")]
        extern static void SmoothingService_KernelSmoothing(ref int functionType, ref int kernel, ref int sample_size, double[] x, double[] pts, double[] cdf);

        [ExcelFunction(Name = "ModelService_Version", Description = "Version Information", Category = "Model Service")]
        public static string PrintVersion()
        {
            string today = DateTime.Now.ToString("yyyy-MM-dd");
            return "Copula Service: Version 1.0.0, Build on " + today;
        }

        [ExcelFunction(Name = "ModelService_DisplayFunctions", Description = "DisplayFunctions", Category = "Model Service")]
        public static object[,] PrintFunctions()
        {
            object[,] retval = new object[8, 2];
            retval[0, 0] = "ModelService_Version"; retval[0, 1] = "Version Information";
            retval[1, 0] = "CopulaService_GaussianCalibration"; retval[1, 1] = "Calibrate Gaussian Copula";
            retval[2, 0] = "CopulaService_StudentCalibration"; retval[2, 1] = "Calibrate Student Copula";
            retval[3, 0] = "CopulaService_ArchimedeanCalibration"; retval[3, 1] = "Calibrate Archimedean Copula";
            retval[4, 0] = "CopulaService_GaussianSimulation"; retval[4, 1] = "Simulate Gaussian Copula";
            retval[5, 0] = "CopulaService_StudentSimulation"; retval[5, 1] = "Simulate Student Copula";
            retval[6, 0] = "CopulaService_ArchimedeanSimulation"; retval[6, 1] = "Simulate Archimedean Copula";
            retval[7, 0] = "SmoothingService_KernelSmoothing"; retval[7, 1] = "Smooth Discrete Data using Kernel Functions";
            return retval;
        }

        [ExcelFunction(Name = "SmoothingService_KernelSmoothing", Description = "Smooth Discrete Data using Kernel Functions", Category = "Kernel Smoothing")]
        public static object[,] KernelSmoothing([ExcelArgument(Name = "functionType", Description = "either cdf or icdf")]string functionType, [ExcelArgument(Name = "kernelName", Description = "Kernel Name")]string kernel, [ExcelArgument(Name = "eval", Description = "Eval Pts")]double[] pts, [ExcelArgument(Name = "data", Description = "Discrete Data")]double[] data)
        {
            int f = functionType == "cdf" ? 0 : 1;
            int k = kernel == "normal" ? 0 : (kernel == "box" ? 1 : (kernel == "triangle" ? 2 : (kernel == "quadratic" ? 3 : 0)));
            int numPts = pts.GetLength(0);
            double[] cdf = new double[numPts];
            SmoothingService_KernelSmoothing(ref f, ref k, ref numPts, pts, data, cdf);
            object[,] retval = new object[numPts + 1, 2];
            retval[0, 0] = "Pts"; retval[0, 1] = "Distribution";
            for(int i = 0; i < numPts; i++)
            {
                retval[i + 1, 0] = pts[i];
                retval[i + 1, 1] = cdf[i];
            }
            return retval;
        }

        [ExcelFunction(Name = "CopulaService_GaussianCalibration", Description = "Calibrate Gaussian Copula", Category = "Copula Calibration")]
        public static object[,] GaussianCalibration([ExcelArgument(Name = "data", Description = "Fitting Data")]object[,] data)
        {
            int sample_size = data.GetLength(0), dim = data.GetLength(1);
            double[] correl = new double[dim * dim], mdata = new double[sample_size*dim];
            int cnt = 0;
            for (int i = 0; i < sample_size; i++)
            {
                for (int j = 0; j < dim; j++)
                {
                    mdata[cnt] = System.Convert.ToDouble(data[i, j]);
                    cnt++;
                }
            }
            object[,] retval = new object[dim, dim + 1];
            CopulaService_GaussianCopulaCalibration(correl, ref dim, ref sample_size, mdata);
            cnt = 0;
            for (int i = 0; i < dim; i++)
            {
                for (int j = 0; j < dim + 1; j++)
                {
                    if (i == 0 && j == 0)
                    {
                        retval[i, j] = "Correlation:";
                    }
                    else if (j == 0)
                    {
                        retval[i, j] = "";
                    }
                    else
                    {
                        retval[i, j] = correl[cnt];
                        cnt++;
                    }
                }
            }
            return retval;
        }

        [ExcelFunction(Name = "CopulaService_StudentCalibration", Description = "Calibrate Student Copula", Category = "Copula Calibration")]
        public static object[,] StudentCalibration([ExcelArgument(Name = "data", Description = "Fitting Data")]object[,] data)
        {
            double degreeOfFreedom = 0.0;
            int sample_size = data.GetLength(0), dim = data.GetLength(1);
            double[] correl = new double[dim * dim], mdata = new double[sample_size * dim];
            int cnt = 0;
            for (int i = 0; i < sample_size; i++)
            {
                for (int j = 0; j < dim; j++)
                {
                    mdata[cnt] = System.Convert.ToDouble(data[i, j]);
                    cnt++;
                }
            }
            object[,] retval = new object[dim + 1, dim + 1];
            CopulaService_StudentCopulaCalibration(ref degreeOfFreedom, correl, ref dim, ref sample_size, mdata);
            cnt = 0;
            for (int i = 0; i < dim + 1; i++)
            {
                for (int j = 0; j < dim + 1; j++)
                {
                    if (i == 0 && j == 0)
                    {
                        retval[i, j] = "DegreeOfFreedom:";
                    }
                    else if (i == 0 && j == 1)
                    {
                        retval[i, j] = degreeOfFreedom;
                    }
                    else if (i == 1 && j == 0)
                    {
                        retval[i, j] = "Correlation";
                    }
                    else if (i == 0 || j == 0)
                    {
                        retval[i, j] = "";
                    }
                    else
                    {
                        retval[i, j] = correl[cnt];
                        cnt++;
                    }
                }
            }
            return retval;
        }

        [ExcelFunction(Name = "CopulaService_ArchimedeanCalibration", Description = "Calibrate Archimedean Copula", Category = "Copula Calibration")]
        public static object[,] ArchimedeanCalibration([ExcelArgument(Name = "type", Description = "Copula Type")]string type, [ExcelArgument(Name = "data", Description = "Fitting Data")]object[,] data)
        {
            double theta = 0.0;
            int t = type == "amh" ? 0 : (type == "clayton" ? 1 : (type == "frank" ? 2 : (type == "gumbel" ? 3 : (type == "joe" ? 4 : 1))));
            int sample_size = data.GetLength(0), dim = data.GetLength(1);
            double[] correl = new double[dim * dim], mdata = new double[sample_size * dim];
            int cnt = 0;
            for (int i = 0; i < sample_size; i++)
            {
                for (int j = 0; j < dim; j++)
                {
                    mdata[cnt] = System.Convert.ToDouble(data[i, j]);
                    cnt++;
                }
            }
            object[,] retval = new object[2, 2];
            CopulaService_ArchimedeanCopulaCalibration(ref t, ref theta, ref dim, ref sample_size, mdata);
            retval[0, 0] = "CopulaType:";  retval[0, 1] = type;
            retval[1, 0] = "Theta:"; retval[1, 1] = theta;
            return retval;
        }

        [ExcelFunction(Name = "CopulaService_GaussianSimulation", Description = "Simulate Gaussian Copula", Category = "Copula Simulation")]
        public static object[,] GaussianSimulation([ExcelArgument(Name = "correlation", Description = "Copula Correlation")]object[,] correl, [ExcelArgument(Name = "size", Description = "Simulation Sample Size")]int sample_size)
        {
            if (correl.GetLength(0) != correl.GetLength(1))
            {
                return new object[,] { { "Dimension error." } };
            }
            int dim = correl.GetLength(0);
            object[,] retval = new object[sample_size + 1, correl.GetLength(0)];
            double[] output = new double[sample_size * correl.GetLength(0)];
            double[] correlation = new double[correl.GetLength(0) * correl.GetLength(1)];
            int cnt = 0;
            for(int i = 0; i < correl.GetLength(0); i++)
            {
                for(int j = 0; j < correl.GetLength(1); j++)
                {
                    correlation[cnt] = System.Convert.ToDouble(correl[i, j]);
                    cnt++;
                }
            }
            CopulaService_GaussianCopulaSimulation(correlation, ref dim, ref sample_size, output);
            cnt = 0;
            for (int i = 0; i < sample_size + 1; i++)
            {
                for (int j = 0; j < correl.GetLength(0); j++)
                {
                    if (i == 0)
                    {
                        retval[i, j] = "sample " + (j + 1).ToString();
                    }
                    else
                    {
                        retval[i, j] = output[cnt];
                        cnt++;
                    }
                }
            }
            return retval;
        }

        [ExcelFunction(Name = "CopulaService_StudentSimulation", Description = "Simulate Student Copula", Category = "Copula Simulation")]
        public static object[,] StudentSimulation([ExcelArgument(Name = "correlation", Description = "Copula Correlation")]object[,] correl, [ExcelArgument(Name = "degreeOfFreedom", Description = "Copula DegreeOfFreedom")]double degreeOfFreedom, [ExcelArgument(Name = "size", Description = "Simulation Sample Size")]int sample_size)
        {
            if (correl.GetLength(0) != correl.GetLength(1))
            {
                return new object[,] { { "Dimension error." } };
            }
            int dim = correl.GetLength(0);
            object[,] retval = new object[sample_size + 1, correl.GetLength(0)];
            double[] output = new double[sample_size * correl.GetLength(0)];
            double[] correlation = new double[correl.GetLength(0) * correl.GetLength(1)];
            int cnt = 0;
            for (int i = 0; i < correl.GetLength(0); i++)
            {
                for (int j = 0; j < correl.GetLength(1); j++)
                {
                    correlation[cnt] = System.Convert.ToDouble(correl[i, j]);
                    cnt++;
                }
            }
            CopulaService_StudentCopulaSimulation(ref degreeOfFreedom, correlation, ref dim, ref sample_size, output);
            cnt = 0;
            for (int i = 0; i < sample_size + 1; i++)
            {
                for (int j = 0; j < correl.GetLength(0); j++)
                {
                    if (i == 0)
                    {
                        retval[i, j] = "sample " + (j + 1).ToString();
                    }
                    else
                    {
                        retval[i, j] = output[cnt];
                        cnt++;
                    }
                }
            }
            return retval;
        }

        [ExcelFunction(Name = "CopulaService_ArchimedeanSimulation", Description = "Simulate Archimedean Copula", Category = "Copula Simulation")]
        public static object[,] ArchimedeanSimulation([ExcelArgument(Name = "type", Description = "Copula Type")]string type, [ExcelArgument(Name = "theta", Description = "Copula Theta")]double theta, [ExcelArgument(Name = "dimension", Description = "Simulation Dimension")]int dim, [ExcelArgument(Name = "size", Description = "Simulation Sample Size")]int sample_size)
        {
            object[,] retval = new object[sample_size + 1, dim];
            double[] output = new double[sample_size * dim];
            int t = type == "amh" ? 0 : (type == "clayton" ? 1 : (type == "frank" ? 2 : (type == "gumbel" ? 3 : (type == "joe" ? 4 : 1))));
            CopulaService_ArchimedeanCopulaSimulation(ref t, ref theta, ref dim, ref sample_size, output);
            int cnt = 0;
            for (int i = 0; i < sample_size + 1; i++)
            {
                for (int j = 0; j < dim; j++)
                {
                    if (i == 0)
                    {
                        retval[i, j] = "sample " + (j + 1).ToString();
                    }
                    else
                    {
                        retval[i, j] = output[cnt];
                        cnt++;
                    }
                }
            }
            return retval;
        }
    }
}
