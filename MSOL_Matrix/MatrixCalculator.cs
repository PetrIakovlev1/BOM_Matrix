using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;

namespace MSOL_Matrix
{
    public static class MatrixCalculator
    {
        public static DataTablesModel calculateRawMatrix(DataTable inputDt)
        {
            DataTablesModel dataTablesModel = new DataTablesModel();
            dataTablesModel.MatrixDt = new DataTable();

            string orderColumnName = "Master-order";
            string itemColumnName = "Material reference";
            string qtyColumnName = "Quantity";

            string[] columnsForItems = { orderColumnName, itemColumnName, qtyColumnName };
            string[] columnsForOrders = { orderColumnName, qtyColumnName };

            List<MatrixModel> matrixModelList = GetMatrixModels(inputDt, columnsForItems);

            string[] ordersList = matrixModelList.Where(x => x.Quantity > 0).Select(x => x.MasterOrder.ToUpper().Trim()).Distinct().ToArray();

            addColumnsToOutputDt(dataTablesModel.MatrixDt, ordersList, orderColumnName);

            dataTablesModel.MatrixDt0Only = dataTablesModel.MatrixDt.Clone();
            dataTablesModel.MatrixDt25 = dataTablesModel.MatrixDt.Clone();
            dataTablesModel.MatrixDt50 = dataTablesModel.MatrixDt.Clone();
            dataTablesModel.MatrixDt75 = dataTablesModel.MatrixDt.Clone();
            dataTablesModel.MatrixDt100 = dataTablesModel.MatrixDt.Clone();
            dataTablesModel.MatrixDt100Only = dataTablesModel.MatrixDt.Clone();

            double[,] simArray = new double[ordersList.Length, ordersList.Length];

            initArray(ref simArray);

            double[,] simArray0Only = (double[,])simArray.Clone();
            double[,] simArray25 = (double[,])simArray.Clone();
            double[,] simArray50 = (double[,])simArray.Clone();
            double[,] simArray75 = (double[,])simArray.Clone();
            double[,] simArray100 = (double[,])simArray.Clone();
            double[,] simArray100Only = (double[,])simArray.Clone();


            int parts = Environment.ProcessorCount;
            int partSize = ordersList.Count() / parts;

            Parallel.For(0, parts, new ParallelOptions(), (k) =>
            {
                for (int i = k * partSize; i < (k + 1) * partSize; i++)
                {
                    for (int j = i; j < ordersList.Length; j++)
                    {
                        double similarity = GetItemListsSimilarity(matrixModelList, ordersList[i], ordersList[j]);
                        simArray[i, j] = similarity;

                        if (similarity == 0.0)
                            simArray0Only[i, j] = similarity;
                        else if (similarity <= 0.25)
                            simArray25[i, j] = similarity;
                        else if (similarity <= 0.5)
                            simArray50[i, j] = similarity;
                        else if (similarity <= 0.75)
                            simArray75[i, j] = similarity;
                        else if (similarity < 1)
                            simArray100[i, j] = similarity;
                        else if (similarity == 1)
                            simArray100Only[i, j] = similarity;
                    }
                }
            });

            dataFromArrayToDt(dataTablesModel.MatrixDt, ordersList, simArray);

            dataFromArrayToDt(dataTablesModel.MatrixDt0Only, ordersList, simArray0Only);
            CleanDt(dataTablesModel.MatrixDt0Only);

            dataFromArrayToDt(dataTablesModel.MatrixDt25, ordersList, simArray25);
            CleanDt(dataTablesModel.MatrixDt25);

            dataFromArrayToDt(dataTablesModel.MatrixDt50, ordersList, simArray50);
            CleanDt(dataTablesModel.MatrixDt50);

            dataFromArrayToDt(dataTablesModel.MatrixDt75, ordersList, simArray75);
            CleanDt(dataTablesModel.MatrixDt75);

            dataFromArrayToDt(dataTablesModel.MatrixDt100, ordersList, simArray100);
            CleanDt(dataTablesModel.MatrixDt100);

            dataFromArrayToDt(dataTablesModel.MatrixDt100Only, ordersList, simArray100Only);
            CleanColumnsDtFor100Only(dataTablesModel.MatrixDt100Only);
            CleanDt(dataTablesModel.MatrixDt100Only);
            return dataTablesModel;
        }

        private static void dataFromArrayToDt(DataTable outputDt, string[] ordersList, double[,] simArray)
        {
            for (int i = 0; i < ordersList.Length; i++)
            {
                DataRow dr = outputDt.NewRow();

                double[] simData = matrixRowToVector(simArray, i);

                var objData = simData.Cast<object>().ToArray();
                objData[0] = ordersList[i];

                dr.ItemArray = objData;
                outputDt.Rows.Add(dr);
            }
        }

        static double[] matrixRowToVector(double[,] simArray, int row)
        {
            double[] simData = new double[simArray.GetLength(0) + 1];

            for (int j = 0; j < simArray.GetLength(1); j++)
            {
                simData[j + 1] = simArray[row, j];
            }

            return simData;
        }

        private static void addColumnsToOutputDt(DataTable outputDt, string[] ordersList, string orderColumnName)
        {
            outputDt.Columns.Add(orderColumnName, typeof(string));

            for (int i = 0; i < ordersList.Length; i++)
            {
                outputDt.Columns.Add(ordersList[i], typeof(double));
            }
        }

        static double GetItemListsSimilarity(List<MatrixModel> matrixModels, string orderFromColumn, string orderFromRow)
        {
            double ItemListsSimilarity = 1.0;

            string[] ItemListFromColumn = GetItemsFromLIst(matrixModels, orderFromColumn);
            string[] ItemListFromRow = GetItemsFromLIst(matrixModels, orderFromRow);

            List<string> diffItemsList = ItemListFromColumn.Except(ItemListFromRow).ToList();
            diffItemsList.AddRange(ItemListFromRow.Except(ItemListFromColumn).ToList());

            if (diffItemsList.Count() > 0)
                ItemListsSimilarity = Math.Round(1 - ((double)diffItemsList.Count() / ((double)ItemListFromColumn.Count() + (double)ItemListFromRow.Count())), 2);

            return ItemListsSimilarity;
        }


        static string[] GetItemsFromLIst(List<MatrixModel> matrixModels, string orderFrom)
        {
            List<MatrixModel> dataRows = matrixModels.Where(x => x.MasterOrder == orderFrom).ToList();
            string[] ItemListFrom = new string[dataRows.Count()];

            int idx = 0;
            int maxJ = dataRows.Count();

            for (int j = 0; j < maxJ; j++)
            {
                var item = dataRows[j];

                if (item.Quantity == 0)
                    item.Quantity = 1;

                for (int i = 0; i < item.Quantity; i++)
                {
                    if (item.Quantity > 1 && i == 0)
                        Array.Resize(ref ItemListFrom, ItemListFrom.Length + (item.Quantity - 1));

                    ItemListFrom.SetValue($"{item.MaterialReference}_{i}", idx);
                    idx++;
                }
            }

            return ItemListFrom;
        }


        static List<MatrixModel> GetMatrixModels(DataTable inputDt, string[] columnsForItems)
        {
            List<MatrixModel> matrixModelList = new List<MatrixModel>();

            foreach (DataRow dr in inputDt.Rows)
            {
                MatrixModel matrixModel = new MatrixModel();

                matrixModel.MasterOrder = ConvertFunctions.GetConvertedToString(dr[columnsForItems[0]]).ToUpper().Trim();
                matrixModel.MaterialReference = ConvertFunctions.GetConvertedToString(dr[columnsForItems[1]]).ToUpper().Trim();
                matrixModel.Quantity = ConvertFunctions.GetConvertedToIntValue(dr[columnsForItems[2]]);

                var Exist = matrixModelList.Where(x => x.MasterOrder == matrixModel.MasterOrder && x.MaterialReference == matrixModel.MaterialReference).FirstOrDefault();

                if (Exist != null)
                {
                    int idx = matrixModelList.IndexOf(Exist);
                    matrixModelList[idx].Quantity += matrixModel.Quantity;
                }
                else
                {
                    matrixModelList.Add(matrixModel);
                }

            }


            return matrixModelList;
        }

        static void initArray(ref double[,] simArray)
        {
            for (int i = 0; i < simArray.GetLength(0); i++)
            {
                for (int j = 0; j < simArray.GetLength(1); j++)
                {
                    simArray[i, j] = -1;
                }
            }
        }

        static void CleanDt(DataTable dt)
        {

            List<int> RowsIndexToDelList = new List<int>();
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                object[] itemArray = dt.Rows[i].ItemArray;
                itemArray[0] = -1;

                if (itemArray.All(o => ConvertFunctions.GetConvertedToDblValue(o) == -1))
                    RowsIndexToDelList.Add(i);

            }

            for (int i = RowsIndexToDelList.Count - 1; i >= 0; i--)
            {
                dt.Rows.RemoveAt(RowsIndexToDelList[i]);
            }

            List<string> columnsToRemove = new List<string>();
            foreach (DataColumn dc in dt.Columns)
            {
                int idx = dt.Columns.IndexOf(dc);

                if (idx > 0)
                    if (dt.AsEnumerable().All(x => x.Field<double>(dc.ColumnName) < 0))
                        columnsToRemove.Add(dc.ColumnName);
            }

            foreach (string c in columnsToRemove)
            {
                dt.Columns.Remove(c);
            }
        }

        static void CleanColumnsDtFor100Only(DataTable dt)
        {

            List<string> columnsToRemove = new List<string>();
            foreach (DataColumn dc in dt.Columns)
            {
                int idx = dt.Columns.IndexOf(dc);

                if (idx > 0)
                {
                    var sum = dt.AsEnumerable().Where(x => x.Field<double>(dc.ColumnName)==1).Sum(x => x.Field<double>(dc.ColumnName));
                    if (sum == 1)
                        columnsToRemove.Add(dc.ColumnName);
                }
            }

            foreach (string c in columnsToRemove)
            {
                dt.Columns.Remove(c);
            }
        }
    }
}
