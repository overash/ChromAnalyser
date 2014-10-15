using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Xml.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
public class SerializationAdapter
{
    BinaryFormatter formatter;
    string path;
    ItemsState state;
    CheckedListBox owner;
     
    public SerializationAdapter(string path, CheckedListBox CheckBox)
    {
        this.state = new ItemsState(CheckBox);
        this.path = path;
        this.formatter = new BinaryFormatter();
        this.owner = CheckBox;
    }

    public void Serialize()
    {
        using (FileStream fs = new FileStream(path, FileMode.OpenOrCreate))
        {
            formatter.Serialize(fs, state);
        }
    }

    public void Deserialize()
    {
        using (FileStream fs = new FileStream(path, FileMode.OpenOrCreate))
        {
            state = (ItemsState)formatter.Deserialize(fs);
        }
        owner.Items.Clear();
        for (int i = 0; i < state.state.Count; i++)
        {
            string itemName = state.state.Keys.ToArray()[i].ToString();
            owner.Items.Add(itemName);
            owner.SetItemChecked(i, state.state[itemName]);
        }
    }

}

[Serializable]
public class ItemsState
{
    public ItemsState()
    {

    }

    public Dictionary<string, bool> state;
    public ItemsState(CheckedListBox ListBox)
    {
        state = new Dictionary<string, bool>();

        for (int i = 0; i < ListBox.Items.Count; i++)
        {
            state.Add(ListBox.Items[i].ToString(), ListBox.GetItemChecked(i));
        }
    }
   
}
