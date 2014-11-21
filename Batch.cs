using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;

public class BatchProcess
{
    public struct FieldValue
    {
        public string Name;
        public string Value;

        public FieldValue(string Name, string Value)
        {
            this.Name = Name;
            this.Value = Value;
        }
    }

    public static enum OnErrorAction
    {
        Return, Continue
    }

    /* /////////////// PUBLIC /////////////////// */

    #region Public
    public OnErrorAction OnError { get; set; }

    public BatchProcess(SPWeb Web)
    {
        _web = Web;
        _methodBuilders = new List<StringBuilder>();
        _addNewMethodBuilder();
        _totalMethodsCount = 0;
        OnError = OnErrorAction.Return;
    }

    public string Run()
    {
        StringBuilder result = new StringBuilder();

        foreach (var methodBuilder in _methodBuilders)
        {
            string Batch = String.Format(BATCH_FORMAT, OnError.ToString("g"), methodBuilder.ToString());
            result.AppendLine(_web.ProcessBatchData(Batch));
        }

        return result.ToString();
    }

    public void AddItem(SPList List, params KeyValuePair<string, string>[] FieldsValues)
    {
        AddItem(List, FieldsValues.Select(fieldValue => new FieldValue(fieldValue.Key, fieldValue.Value)).ToArray());
    }

    public void AddItem(SPList List, params FieldValue[] FieldsValues)
    {
        _addAction(List, "Save", "New", FieldsValues);
    }

    public void EditItem(SPList List, int ID, params KeyValuePair<string, string>[] FieldsValues)
    {
        EditItem(List, ID, FieldsValues.Select(fieldValue => new FieldValue(fieldValue.Key, fieldValue.Value)).ToArray());
    }

    public void EditItem(SPList List, int ID, params FieldValue[] FieldsValues)
    {
        _addAction(List, "Save", ID.ToString(), FieldsValues);
    }

    public void DeleteItem(SPList List, int ID)
    {
        _addAction(List, "Delete", ID.ToString());
    }

    #endregion

    /* /////////////// PRIVATE /////////////////// */

    #region Private
    const string BATCH_FORMAT = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><ows:Batch OnError=\"{0}\">{1}</ows:Batch>";
    const string METHOD_FORMAT =
    "<Method ID=\"{0}\">" +
        "<SetList>{1}</SetList>" +
        "<SetVar Name=\"Cmd\">{2}</SetVar>" +
        "<SetVar Name=\"ID\">{3}</SetVar>" +
        "{4}" +
    "</Method>";
    const string SET_VAR_FORMAT = "<SetVar Name=\"urn:schemas-microsoft-com:office:office#{0}\">{1}</SetVar>";
    const int BATCH_LIMIT = 500;

    private SPWeb _web;
    private List<StringBuilder> _methodBuilders;
    private StringBuilder _currentMethodBuilder;
    private int _currentMethodsCount;
    private int _totalMethodsCount;

    private void _addNewMethodBuilder()
    {
        _methodBuilders.Add(new StringBuilder());

        _currentMethodBuilder = _methodBuilders[_methodBuilders.Count - 1];

        _currentMethodsCount = 0;
    }

    private string _getSetVar(params FieldValue[] FieldsValues)
    {
        StringBuilder SetVarBuilder = new StringBuilder();

        foreach (var fieldValue in FieldsValues)
        {
            SetVarBuilder.AppendFormat(SET_VAR_FORMAT, fieldValue.Name, fieldValue.Value);
        }

        return SetVarBuilder.ToString();
    }

    private void _addAction(SPList List, string Action, string ID, params FieldValue[] FieldsValues)
    {
        if (_currentMethodsCount >= BATCH_LIMIT)
        {
            _addNewMethodBuilder();
        }

        _currentMethodBuilder.AppendFormat(
            METHOD_FORMAT,
            _totalMethodsCount++,
            List.ID,
            Action,
            ID,
            _getSetVar(FieldsValues)
        );

        ++_currentMethodsCount;
    }
    #endregion

    /* ///////////////////////////////////////// */
}