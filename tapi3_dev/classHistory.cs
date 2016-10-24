using System;
using System.Collections.Generic;
using System.Text;

namespace tapi3_dev
{
  class classHistory
  {
    private string direction;
    private string fromName;
    private string fromNumber;
    private string toName;
    private string toNumber; 
    private DateTime startTime;
    private int callID;

    public string FromName
    {
      get { return (fromName ?? "").ToString(); }
      set
      {
        Encoding wind1252 = Encoding.GetEncoding(1252);
        Encoding utf8 = Encoding.UTF8;
        byte[] wind1252Bytes = wind1252.GetBytes(value);
        string wind1252String = wind1252.GetString(wind1252Bytes);
        byte[] utf8Bytes = Encoding.Convert(wind1252, utf8, wind1252Bytes);
        fromName = Encoding.UTF8.GetString(utf8Bytes);
      }
    }

    public string FromNumber
    {
      get { return (fromNumber ?? "").ToString(); }
      set { fromNumber = value; }
    }

    public string ToName
    {
      get { return (toName ?? "").ToString();}
      set { toName = value;
      }
    }

    public string ToNumber
    {
      get { return (toNumber ?? "").ToString(); }
      set { toNumber = value;}
    }

    public DateTime StartTime
    {
      get { return startTime;}
      set { startTime = value;}
    }

    public string Direction
    {
      get { return (direction ?? "").ToString(); }
      set { direction = value; }
    }

    public int CallID
    {
      get { return callID; }
      set { callID = value; }
    }

    public void addHistory(string fromName, string fromNumber, string toName, string toNumber, DateTime startTime, string direction, int callID)
    {
      this.FromName = fromName;
      this.FromNumber = fromNumber;
      this.ToName = toName;
      this.ToNumber = toNumber;
      this.StartTime = startTime;
      this.Direction = direction;
      this.CallID = callID;
    }

    public void Clear()
    {
      this.FromName = "";
      this.FromNumber = "";
      this.ToName = "";
      this.ToNumber = "";
      this.Direction = "";
      this.CallID = 0;
    }
  }
}
