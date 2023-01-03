using System.Data;
using System.Xml.Serialization;
using FastMember;
using OpenXmlClient.Classes.Models;
using OpenXmlClient.FormatSettings;
using OpenXmlClient.Models.Table;

#pragma warning disable CS8625

namespace OpenXmlClient.Classes;

    public static class TableRowDataTableSerializer
    {
        /// <summary>
        /// Method for serializing row data to xml DataTable
        /// </summary>
        /// <param name="tableRowFillModel"></param>
        /// <param name="tableValueOutputFormat"></param>
        /// <returns></returns>
        public static RowsRenderPayload FillRows(TableRowFillModel tableRowFillModel, 
            TableValueOutputFormat tableValueOutputFormat = null)
        {
            var  serializer = new XmlSerializer(typeof(DataTable));
            var payload = new RowsRenderPayload();
            payload.ColumnNames = tableRowFillModel.ColumnNames;
            payload.TableValueOutputFormat = tableRowFillModel.TableValueOutputFormat;
            SetTableOutputFormat(tableValueOutputFormat, payload);
            payload.IsRemoveEntireRowIfNoRecords = tableRowFillModel.RemoveEntireRowIfNoRecords;
            payload.RemoveInnerBordersInColumnAbsolutePositions =
                tableRowFillModel.RemoveInnerBordersInColumnAbsolutePositions;

            if (tableRowFillModel.TableData == null || !tableRowFillModel.TableData.Any())
            {
                return payload;
            }

            if (tableRowFillModel.TableData.Any(p => p == null))
            {
                throw new Exception("LOAN_CORPORATE_PRINT_FORM_SERVICE/ONE_OF_RECORD_HAS_NULL_VALUE");
            }

            // FastMember - читает в DataTable
            var dataTable = new DataTable(Guid.NewGuid().ToString());
            // забираем название класса
            var type = tableRowFillModel.TableData.GetType().GetGenericArguments()[0];
            using (var reader = new ObjectReader(type,
                       tableRowFillModel.TableData.ToArray()))
            {
                dataTable.Load(reader);
            }
            
            string requestData;
            using (var sw = new StringWriter())
            {
                serializer.Serialize(sw, dataTable);
                requestData = sw.ToString();
            }

            payload.Payload = requestData;
            
            return payload;
            
        }


        /// <summary>
        /// Set table output format
        /// </summary>
        /// <param name="outerTableValueOutputFormat"></param>
        /// <param name="rowsRenderPayload"></param>
        private static void SetTableOutputFormat(TableValueOutputFormat outerTableValueOutputFormat,
            RowsRenderPayload rowsRenderPayload)
        {
            if (rowsRenderPayload.TableValueOutputFormat == null)
            {
                rowsRenderPayload.TableValueOutputFormat = new TableValueOutputFormat();
            }
            if (outerTableValueOutputFormat != null)
            {
                TableValueOutputFormat.SetOutputFormat(outerTableValueOutputFormat,
                      rowsRenderPayload.TableValueOutputFormat);
                 TableValueOutputFormat.SetOutputFormat(WinWordRenderModel.TableValueOutputFormat,
                         rowsRenderPayload.TableValueOutputFormat);
            }
            else 
            {
                TableValueOutputFormat.SetOutputFormat(WinWordRenderModel.TableValueOutputFormat,
                        rowsRenderPayload.TableValueOutputFormat);
            }
        }
        
      
    }