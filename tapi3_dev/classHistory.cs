using System;
using System.Collections.Generic;
using System.Text;

namespace tapi3_dev
{
  class classHistory
  {
    private string fromName;
    private string fromNumber;
    private string toName;
    private string toNumber;
    private DateTime startTime;

    public string FromName
    {
      get { return fromName; }
      set { fromName = value;}
    }

    public string FromNumber
    {
      get { return fromNumber; }
      set { fromNumber = value; }
    }

    public string ToName
    {
      get { return toName;}
      set { toName = value;
      }
    }

    public string ToNumber
    {
      get { return toNumber; }
      set { toNumber = value;}
    }

    public DateTime StartTime
    {
      get { return startTime;}
      set { startTime = value;}
    }

    public void addHistory(string fromName, string fromNumber, string toName, string toNumber, DateTime startTime)
    {
      this.FromName = fromName;
      this.FromNumber = fromNumber;
      this.ToName = toName;
      this.ToNumber = toNumber;
      this.StartTime = startTime;
    }
  }
}
