﻿//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:2.0.50727.5466
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Web.Services;
using System.Web.Services.Protocols;
using System.Xml.Serialization;

// 
// This source code was auto-generated by wsdl, Version=2.0.50727.3038.
// 


/// <remarks/>
[System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.3038")]
[System.Diagnostics.DebuggerStepThroughAttribute()]
[System.ComponentModel.DesignerCategoryAttribute("code")]
[System.Web.Services.WebServiceBindingAttribute(Name="NEPSLoadEquipmentBinding", Namespace="http://www.tm.com.my/hsbb/eai/NEPSLoadEquipment/")]
public partial class NEPSLoadEquipmentService : System.Web.Services.Protocols.SoapHttpClientProtocol {
    
    private System.Threading.SendOrPostCallback NEPSLoadEquipmentOperationCompleted;
    
    /// <remarks/>
    public NEPSLoadEquipmentService() {
        this.Url = "http://10.14.38.171:7001/prj_HsbbEai_Sync_War/ws/NEPSLoadEquipment/";
    }
    
    /// <remarks/>
    public event NEPSLoadEquipmentCompletedEventHandler NEPSLoadEquipmentCompleted;
    
    /// <remarks/>
    [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://www.tm.com.my/hsbb/eai/NEPSLoadEquipment", RequestElementName="NEPSLoadEquipmentRequest", RequestNamespace="http://www.tm.com.my/hsbb/eai/NEPSLoadEquipment/", ResponseNamespace="http://www.tm.com.my/hsbb/eai/NEPSLoadEquipment/", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Wrapped)]
    [return: System.Xml.Serialization.XmlElementAttribute("EquipmentList", Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
    public NEPSLoadEquipmentResponseEquipmentList NEPSLoadEquipment([System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)] string eventName, [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)] NEPSLoadEquipmentRequestSetEquip setEquip, [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)] NEPSLoadEquipmentRequestSetEquipUDA setEquipUDA, [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)] out NEPSLoadEquipmentResponseListEquipUDA ListEquipUDA, [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)] out NEPSLoadEquipmentResponseCallstatus Callstatus) {
        object[] results = this.Invoke("NEPSLoadEquipment", new object[] {
                    eventName,
                    setEquip,
                    setEquipUDA});
        ListEquipUDA = ((NEPSLoadEquipmentResponseListEquipUDA)(results[1]));
        Callstatus = ((NEPSLoadEquipmentResponseCallstatus)(results[2]));
        return ((NEPSLoadEquipmentResponseEquipmentList)(results[0]));
    }
    
    /// <remarks/>
    public System.IAsyncResult BeginNEPSLoadEquipment(string eventName, NEPSLoadEquipmentRequestSetEquip setEquip, NEPSLoadEquipmentRequestSetEquipUDA setEquipUDA, System.AsyncCallback callback, object asyncState) {
        return this.BeginInvoke("NEPSLoadEquipment", new object[] {
                    eventName,
                    setEquip,
                    setEquipUDA}, callback, asyncState);
    }
    
    /// <remarks/>
    public NEPSLoadEquipmentResponseEquipmentList EndNEPSLoadEquipment(System.IAsyncResult asyncResult, out NEPSLoadEquipmentResponseListEquipUDA ListEquipUDA, out NEPSLoadEquipmentResponseCallstatus Callstatus) {
        object[] results = this.EndInvoke(asyncResult);
        ListEquipUDA = ((NEPSLoadEquipmentResponseListEquipUDA)(results[1]));
        Callstatus = ((NEPSLoadEquipmentResponseCallstatus)(results[2]));
        return ((NEPSLoadEquipmentResponseEquipmentList)(results[0]));
    }
    
    /// <remarks/>
    public void NEPSLoadEquipmentAsync(string eventName, NEPSLoadEquipmentRequestSetEquip setEquip, NEPSLoadEquipmentRequestSetEquipUDA setEquipUDA) {
        this.NEPSLoadEquipmentAsync(eventName, setEquip, setEquipUDA, null);
    }
    
    /// <remarks/>
    public void NEPSLoadEquipmentAsync(string eventName, NEPSLoadEquipmentRequestSetEquip setEquip, NEPSLoadEquipmentRequestSetEquipUDA setEquipUDA, object userState) {
        if ((this.NEPSLoadEquipmentOperationCompleted == null)) {
            this.NEPSLoadEquipmentOperationCompleted = new System.Threading.SendOrPostCallback(this.OnNEPSLoadEquipmentOperationCompleted);
        }
        this.InvokeAsync("NEPSLoadEquipment", new object[] {
                    eventName,
                    setEquip,
                    setEquipUDA}, this.NEPSLoadEquipmentOperationCompleted, userState);
    }
    
    private void OnNEPSLoadEquipmentOperationCompleted(object arg) {
        if ((this.NEPSLoadEquipmentCompleted != null)) {
            System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
            this.NEPSLoadEquipmentCompleted(this, new NEPSLoadEquipmentCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
        }
    }
    
    /// <remarks/>
    public new void CancelAsync(object userState) {
        base.CancelAsync(userState);
    }
}

/// <remarks/>
[System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.3038")]
[System.SerializableAttribute()]
[System.Diagnostics.DebuggerStepThroughAttribute()]
[System.ComponentModel.DesignerCategoryAttribute("code")]
[System.Xml.Serialization.XmlTypeAttribute(AnonymousType=true, Namespace="http://www.tm.com.my/hsbb/eai/NEPSLoadEquipment/")]
public partial class NEPSLoadEquipmentRequestSetEquip {
    
    private string equipIDField;
    
    private string mngtIPField;
    
    private string equipCatField;
    
    private string equipVendField;
    
    private string equipModelField;
    
    private string regionField;
    
    private string stateField;
    
    private string exchDescField;
    
    private string siteField;
    
    private string tempNameField;
    
    /// <remarks/>
    [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
    public string EquipID {
        get {
            return this.equipIDField;
        }
        set {
            this.equipIDField = value;
        }
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
    public string MngtIP {
        get {
            return this.mngtIPField;
        }
        set {
            this.mngtIPField = value;
        }
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
    public string EquipCat {
        get {
            return this.equipCatField;
        }
        set {
            this.equipCatField = value;
        }
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
    public string EquipVend {
        get {
            return this.equipVendField;
        }
        set {
            this.equipVendField = value;
        }
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
    public string EquipModel {
        get {
            return this.equipModelField;
        }
        set {
            this.equipModelField = value;
        }
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
    public string Region {
        get {
            return this.regionField;
        }
        set {
            this.regionField = value;
        }
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
    public string State {
        get {
            return this.stateField;
        }
        set {
            this.stateField = value;
        }
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
    public string ExchDesc {
        get {
            return this.exchDescField;
        }
        set {
            this.exchDescField = value;
        }
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
    public string Site {
        get {
            return this.siteField;
        }
        set {
            this.siteField = value;
        }
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
    public string TempName {
        get {
            return this.tempNameField;
        }
        set {
            this.tempNameField = value;
        }
    }
}

/// <remarks/>
[System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.3038")]
[System.SerializableAttribute()]
[System.Diagnostics.DebuggerStepThroughAttribute()]
[System.ComponentModel.DesignerCategoryAttribute("code")]
[System.Xml.Serialization.XmlTypeAttribute(AnonymousType=true, Namespace="http://www.tm.com.my/hsbb/eai/NEPSLoadEquipment/")]
public partial class NEPSLoadEquipmentRequestSetEquipUDA {
    
    private string taggingField;
    
    private string outdoorIndoorTaggingField;
    
    private string dpExtensionFlagField;
    
    private string networkNameField;
    
    /// <remarks/>
    [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
    public string Tagging {
        get {
            return this.taggingField;
        }
        set {
            this.taggingField = value;
        }
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
    public string OutdoorIndoorTagging {
        get {
            return this.outdoorIndoorTaggingField;
        }
        set {
            this.outdoorIndoorTaggingField = value;
        }
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
    public string DpExtensionFlag {
        get {
            return this.dpExtensionFlagField;
        }
        set {
            this.dpExtensionFlagField = value;
        }
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
    public string NetworkName {
        get {
            return this.networkNameField;
        }
        set {
            this.networkNameField = value;
        }
    }
}

/// <remarks/>
[System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.3038")]
[System.SerializableAttribute()]
[System.Diagnostics.DebuggerStepThroughAttribute()]
[System.ComponentModel.DesignerCategoryAttribute("code")]
[System.Xml.Serialization.XmlTypeAttribute(AnonymousType=true, Namespace="http://www.tm.com.my/hsbb/eai/NEPSLoadEquipment/")]
public partial class NEPSLoadEquipmentResponseEquipmentList {
    
    private string equipInstIdField;
    
    private string equipIDField;
    
    private string equipStatusField;
    
    private string equipCatField;
    
    private string equipVendField;
    
    private string equipModelField;
    
    private string mngtIPField;
    
    private string siteField;
    
    private string siteDescField;
    
    private string tempNameField;
    
    /// <remarks/>
    [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
    public string EquipInstId {
        get {
            return this.equipInstIdField;
        }
        set {
            this.equipInstIdField = value;
        }
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
    public string EquipID {
        get {
            return this.equipIDField;
        }
        set {
            this.equipIDField = value;
        }
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
    public string EquipStatus {
        get {
            return this.equipStatusField;
        }
        set {
            this.equipStatusField = value;
        }
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
    public string EquipCat {
        get {
            return this.equipCatField;
        }
        set {
            this.equipCatField = value;
        }
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
    public string EquipVend {
        get {
            return this.equipVendField;
        }
        set {
            this.equipVendField = value;
        }
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
    public string EquipModel {
        get {
            return this.equipModelField;
        }
        set {
            this.equipModelField = value;
        }
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
    public string MngtIP {
        get {
            return this.mngtIPField;
        }
        set {
            this.mngtIPField = value;
        }
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
    public string Site {
        get {
            return this.siteField;
        }
        set {
            this.siteField = value;
        }
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
    public string SiteDesc {
        get {
            return this.siteDescField;
        }
        set {
            this.siteDescField = value;
        }
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
    public string TempName {
        get {
            return this.tempNameField;
        }
        set {
            this.tempNameField = value;
        }
    }
}

/// <remarks/>
[System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.3038")]
[System.SerializableAttribute()]
[System.Diagnostics.DebuggerStepThroughAttribute()]
[System.ComponentModel.DesignerCategoryAttribute("code")]
[System.Xml.Serialization.XmlTypeAttribute(AnonymousType=true, Namespace="http://www.tm.com.my/hsbb/eai/NEPSLoadEquipment/")]
public partial class NEPSLoadEquipmentResponseListEquipUDA {
    
    private string taggingField;
    
    private string outdoorIndoorTaggingField;
    
    private string dpExtensionFlagField;
    
    private string networkNameField;
    
    /// <remarks/>
    [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
    public string Tagging {
        get {
            return this.taggingField;
        }
        set {
            this.taggingField = value;
        }
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
    public string OutdoorIndoorTagging {
        get {
            return this.outdoorIndoorTaggingField;
        }
        set {
            this.outdoorIndoorTaggingField = value;
        }
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
    public string DpExtensionFlag {
        get {
            return this.dpExtensionFlagField;
        }
        set {
            this.dpExtensionFlagField = value;
        }
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
    public string NetworkName {
        get {
            return this.networkNameField;
        }
        set {
            this.networkNameField = value;
        }
    }
}

/// <remarks/>
[System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.3038")]
[System.SerializableAttribute()]
[System.Diagnostics.DebuggerStepThroughAttribute()]
[System.ComponentModel.DesignerCategoryAttribute("code")]
[System.Xml.Serialization.XmlTypeAttribute(AnonymousType=true, Namespace="http://www.tm.com.my/hsbb/eai/NEPSLoadEquipment/")]
public partial class NEPSLoadEquipmentResponseCallstatus {
    
    private string statusField;
    
    private string messageCodeField;
    
    private string messageField;
    
    /// <remarks/>
    [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
    public string Status {
        get {
            return this.statusField;
        }
        set {
            this.statusField = value;
        }
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, DataType="integer")]
    public string MessageCode {
        get {
            return this.messageCodeField;
        }
        set {
            this.messageCodeField = value;
        }
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
    public string Message {
        get {
            return this.messageField;
        }
        set {
            this.messageField = value;
        }
    }
}

/// <remarks/>
[System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.3038")]
public delegate void NEPSLoadEquipmentCompletedEventHandler(object sender, NEPSLoadEquipmentCompletedEventArgs e);

/// <remarks/>
[System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.3038")]
[System.Diagnostics.DebuggerStepThroughAttribute()]
[System.ComponentModel.DesignerCategoryAttribute("code")]
public partial class NEPSLoadEquipmentCompletedEventArgs : System.ComponentModel.AsyncCompletedEventArgs {
    
    private object[] results;
    
    internal NEPSLoadEquipmentCompletedEventArgs(object[] results, System.Exception exception, bool cancelled, object userState) : 
            base(exception, cancelled, userState) {
        this.results = results;
    }
    
    /// <remarks/>
    public NEPSLoadEquipmentResponseEquipmentList Result {
        get {
            this.RaiseExceptionIfNecessary();
            return ((NEPSLoadEquipmentResponseEquipmentList)(this.results[0]));
        }
    }
    
    /// <remarks/>
    public NEPSLoadEquipmentResponseListEquipUDA ListEquipUDA {
        get {
            this.RaiseExceptionIfNecessary();
            return ((NEPSLoadEquipmentResponseListEquipUDA)(this.results[1]));
        }
    }
    
    /// <remarks/>
    public NEPSLoadEquipmentResponseCallstatus Callstatus {
        get {
            this.RaiseExceptionIfNecessary();
            return ((NEPSLoadEquipmentResponseCallstatus)(this.results[2]));
        }
    }
}
