using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Collections.Specialized;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
namespace ExcelAddIn1
{
    public partial class Anonymous : INotifyPropertyChanged
    {
        private ObservableCollection<Anonymous> _value;
        [JsonProperty("value", Required = Required.Default)]
        public ObservableCollection<Anonymous> Value
        {
            get { return _value; }
            set
            {
                if (_value != value)
                {
                    _value = value;
                    RaisePropertyChanged();
                }
            }
        }
        public event PropertyChangedEventHandler PropertyChanged;
        public string ToJson()
        {
            return JsonConvert.SerializeObject(this);
        }
        public static Anonymous FromJson(string data)
        {
            return JsonConvert.DeserializeObject<Anonymous>(data);
        }
        protected virtual void RaisePropertyChanged([CallerMemberName] string propertyName = null)
        {
            var handler = PropertyChanged;
            if (handler != null)
                handler(this, new PropertyChangedEventArgs(propertyName));
        }
    }
    public partial class Anonymous1 : INotifyPropertyChanged
    {
        private string _id;
        private string _name;
        private properties _properties;
        private string _type;
        [JsonProperty("id", Required = Required.Default)]
        public string Id
        {
            get { return _id; }
            set
            {
                if (_id != value)
                {
                    _id = value;
                    RaisePropertyChanged();
                }
            }
        }
        [JsonProperty("name", Required = Required.Default)]
        public string Name
        {
            get { return _name; }
            set
            {
                if (_name != value)
                {
                    _name = value;
                    RaisePropertyChanged();
                }
            }
        }
        [JsonProperty("properties", Required = Required.Default)]
        public properties Properties
        {
            get { return _properties; }
            set
            {
                if (_properties != value)
                {
                    _properties = value;
                    RaisePropertyChanged();
                }
            }
        }
        [JsonProperty("type", Required = Required.Default)]
        public string Type
        {
            get { return _type; }
            set
            {
                if (_type != value)
                {
                    _type = value;
                    RaisePropertyChanged();
                }
            }
        }
        public event PropertyChangedEventHandler PropertyChanged;
        public string ToJson()
        {
            return JsonConvert.SerializeObject(this);
        }
        public static Anonymous1 FromJson(string data)
        {
            return JsonConvert.DeserializeObject<Anonymous1>(data);
        }
        protected virtual void RaisePropertyChanged([CallerMemberName] string propertyName = null)
        {
            var handler = PropertyChanged;
            if (handler != null)
                handler(this, new PropertyChangedEventArgs(propertyName));
        }
    }
    public partial class Anonymous2 : INotifyPropertyChanged
    {
        private string _id;
        private string _name;
        private properties _properties;
        private string _type;
        [JsonProperty("id", Required = Required.Default)]
        public string Id
        {
            get { return _id; }
            set
            {
                if (_id != value)
                {
                    _id = value;
                    RaisePropertyChanged();
                }
            }
        }
        [JsonProperty("name", Required = Required.Default)]
        public string Name
        {
            get { return _name; }
            set
            {
                if (_name != value)
                {
                    _name = value;
                    RaisePropertyChanged();
                }
            }
        }
        [JsonProperty("properties", Required = Required.Default)]
        public properties Properties
        {
            get { return _properties; }
            set
            {
                if (_properties != value)
                {
                    _properties = value;
                    RaisePropertyChanged();
                }
            }
        }
        [JsonProperty("type", Required = Required.Default)]
        public string Type
        {
            get { return _type; }
            set
            {
                if (_type != value)
                {
                    _type = value;
                    RaisePropertyChanged();
                }
            }
        }
        public event PropertyChangedEventHandler PropertyChanged;
        public string ToJson()
        {
            return JsonConvert.SerializeObject(this);
        }
        public static Anonymous2 FromJson(string data)
        {
            return JsonConvert.DeserializeObject<Anonymous2>(data);
        }
        protected virtual void RaisePropertyChanged([CallerMemberName] string propertyName = null)
        {
            var handler = PropertyChanged;
            if (handler != null)
                handler(this, new PropertyChangedEventArgs(propertyName));
        }
    }
    public partial class properties : INotifyPropertyChanged
    {
        private object _infoFields;
        private string _instanceData;
        private string _meterCategory;
        private string _meterId;
        private string _meterName;
        private string _meterSubCategory;
        private decimal _quantity;
        private string _subscriptionId;
        private string _unit;
        private string _usageEndTime;
        private string _usageStartTime;
        [JsonProperty("infoFields", Required = Required.Default)]
        public object InfoFields
        {
            get { return _infoFields; }
            set
            {
                if (_infoFields != value)
                {
                    _infoFields = value;
                    RaisePropertyChanged();
                }
            }
        }
        [JsonProperty("instanceData", Required = Required.Default)]
        public string InstanceData
        {
            get { return _instanceData; }
            set
            {
                if (_instanceData != value)
                {
                    _instanceData = value;
                    RaisePropertyChanged();
                }
            }
        }
        [JsonProperty("meterCategory", Required = Required.Default)]
        public string MeterCategory
        {
            get { return _meterCategory; }
            set
            {
                if (_meterCategory != value)
                {
                    _meterCategory = value;
                    RaisePropertyChanged();
                }
            }
        }
        [JsonProperty("meterId", Required = Required.Default)]
        public string MeterId
        {
            get { return _meterId; }
            set
            {
                if (_meterId != value)
                {
                    _meterId = value;
                    RaisePropertyChanged();
                }
            }
        }
        [JsonProperty("meterName", Required = Required.Default)]
        public string MeterName
        {
            get { return _meterName; }
            set
            {
                if (_meterName != value)
                {
                    _meterName = value;
                    RaisePropertyChanged();
                }
            }
        }
        [JsonProperty("meterSubCategory", Required = Required.Default)]
        public string MeterSubCategory
        {
            get { return _meterSubCategory; }
            set
            {
                if (_meterSubCategory != value)
                {
                    _meterSubCategory = value;
                    RaisePropertyChanged();
                }
            }
        }
        [JsonProperty("quantity", Required = Required.Default)]
        public decimal Quantity
        {
            get { return _quantity; }
            set
            {
                if (_quantity != value)
                {
                    _quantity = value;
                    RaisePropertyChanged();
                }
            }
        }
        [JsonProperty("subscriptionId", Required = Required.Default)]
        public string SubscriptionId
        {
            get { return _subscriptionId; }
            set
            {
                if (_subscriptionId != value)
                {
                    _subscriptionId = value;
                    RaisePropertyChanged();
                }
            }
        }
        [JsonProperty("unit", Required = Required.Default)]
        public string Unit
        {
            get { return _unit; }
            set
            {
                if (_unit != value)
                {
                    _unit = value;
                    RaisePropertyChanged();
                }
            }
        }
        [JsonProperty("usageEndTime", Required = Required.Default)]
        public string UsageEndTime
        {
            get { return _usageEndTime; }
            set
            {
                if (_usageEndTime != value)
                {
                    _usageEndTime = value;
                    RaisePropertyChanged();
                }
            }
        }
        [JsonProperty("usageStartTime", Required = Required.Default)]
        public string UsageStartTime
        {
            get { return _usageStartTime; }
            set
            {
                if (_usageStartTime != value)
                {
                    _usageStartTime = value;
                    RaisePropertyChanged();
                }
            }
        }
        public event PropertyChangedEventHandler PropertyChanged;
        public string ToJson()
        {
            return JsonConvert.SerializeObject(this);
        }
        public static properties FromJson(string data)
        {
            return JsonConvert.DeserializeObject<properties>(data);
        }
        protected virtual void RaisePropertyChanged([CallerMemberName] string propertyName = null)
        {
            var handler = PropertyChanged;
            if (handler != null)
                handler(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
