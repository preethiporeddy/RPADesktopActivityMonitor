using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace RPADesktopActivityMonitor
{
    public static class ConversionScript
    {
        public static ConversionResult PrepareRpaPayload(string desktopRecord, string webRecord)
        {
            if (!(File.Exists(desktopRecord) && File.Exists(webRecord))) return new ConversionResult { Status = false, Data = "" };
            dynamic desktopObject, webObject;
            string line;
            using (StreamReader fs = new StreamReader(new FileStream(desktopRecord, FileMode.Open)))
            {
                line = "";
                while (!fs.EndOfStream)
                {
                    line += fs.ReadLine();
                }
            }
            desktopObject = JsonConvert.DeserializeObject(line);

            using (StreamReader fs = new StreamReader(new FileStream(webRecord, FileMode.Open)))
            {
                line = "";
                while (!fs.EndOfStream)
                {
                    line += fs.ReadLine();
                }
            }
            webObject = JsonConvert.DeserializeObject(line);
            if (webObject != null)
            {
                foreach (var obj in webObject) desktopObject.Add(obj);
            }

            var masterObject = ((JArray)desktopObject).OrderBy(obj => (long)obj["timeStamp"]);

            /*var jsonTree0 = new Dictionary<string, dynamic>();
            var jsonTree1 = new Dictionary<string, dynamic>();
            var jsonTree2 = new Dictionary<string, dynamic>();

            var jsonArrayString = new List<string>();

            jsonTree0.Add("id", "");
            jsonTree0.Add("actions", jsonTree1);

            var action_present = false;
            var parent_id = Guid.NewGuid();

            jsonTree0["id"] = Guid.NewGuid();

            jsonTree1.Add(parent_id.ToString(), jsonTree2);

            jsonTree2.Add("id", parent_id);
            jsonTree2.Add("type", "ADD_SEQUENCE");
            jsonTree2.Add("label", "Sequence");
            jsonTree2.Add("key", "sequence");
            jsonTree2.Add("title", "Stage-1");
            jsonTree2.Add("allowNesting", true);
            jsonTree2.Add("expanded", true);
            jsonTree2.Add("subActions", jsonArrayString);
            jsonTree2.Add("instruction", "// Stage-1 starts here");

            foreach (var item in masterObject)
            {
                var id = Guid.NewGuid();
                var jsonActionArrayString = new List<string>();
                jsonTree2 = new Dictionary<string, dynamic>();

                switch ((string)item["actionType"])
                {
                    case "click":
                        action_present = true;
                        jsonTree1.Add(id.ToString(), jsonTree2);

                        jsonTree2.Add("id", id);
                        jsonTree2.Add("type", "CLICK_EL");
                        jsonTree2.Add("parentId", parent_id);
                        jsonTree2.Add("label", "Click");
                        jsonTree2.Add("key", "click");
                        jsonTree2.Add("allowNesting", false);
                        jsonTree2.Add("selector", item["xPath"]);
                        jsonTree2.Add("selectorType", "elSelector");
                        jsonTree2.Add("elSelector", "");
                        jsonTree2.Add("imgSelector", null);
                        jsonTree2.Add("instruction", $"click {item["xPath"]}");
                        jsonTree2.Add("description", $"click {item["xPath"]}");
                        jsonTree2.Add("status", true);
                        jsonTree2.Add("breakpoint", false);
                        break;
                    case "lclick":
                        action_present = true;
                        jsonTree1.Add(id.ToString(), jsonTree2);

                        jsonTree2.Add("id", id);
                        jsonTree2.Add("type", "CLICK_EL");
                        jsonTree2.Add("parentId", parent_id);
                        jsonTree2.Add("label", "Click");
                        jsonTree2.Add("key", "click");
                        jsonTree2.Add("allowNesting", false);
                        jsonTree2.Add("selector", "");
                        jsonTree2.Add("selectorType", "location");
                        jsonTree2.Add("elSelector", "");
                        jsonTree2.Add("imgSelector", $"{item["image"]}");
                        jsonTree2.Add("instruction", $"vision MouseClick, left, {item["coord"]["x"]}, {item["coord"]["y"]}");
                        jsonTree2.Add("description", $"Click left, {item["coord"]["x"]}, {item["coord"]["y"]}");
                        jsonTree2.Add("status", true);
                        jsonTree2.Add("breakpoint", false);
                        jsonTree2.Add("x", item["coord"]["x"]);
                        jsonTree2.Add("y", item["coord"]["y"]);
                        break;
                    /////////////////////////////////////////////////////////////////////////////////
                    //Conflicting format with webrecorder...
                    /////////////////////////////////////////////////////////////////////////////////
                    //case "rclick": 
                    //    action_present = true;
                    //    jsonTree1.Add(id.ToString(), jsonTree2);

                    //    jsonTree2.Add("id", id);
                    //    jsonTree2.Add("type", "CLICK_EL");
                    //    jsonTree2.Add("parentId", parent_id);
                    //    jsonTree2.Add("label", "Click");
                    //    jsonTree2.Add("key", "click");
                    //    jsonTree2.Add("allowNesting", false);
                    //    jsonTree2.Add("selector", "");
                    //    jsonTree2.Add("selectorType", "location");
                    //    jsonTree2.Add("elSelector", "");
                    //    jsonTree2.Add("imgSelector", $"{item["image"]}");
                    //    jsonTree2.Add("instruction", $"vision MouseClick, right, {item["coord"]["x"]}, {item["coord"]["y"]}");
                    //    jsonTree2.Add("description", $"Click right, {item["coord"]["x"]}, {item["coord"]["y"]}");
                    //    jsonTree2.Add("status", true);
                    //    jsonTree2.Add("breakpoint", false);
                    //    jsonTree2.Add("x", item["coord"]["x"]);
                    //    jsonTree2.Add("y", item["coord"]["y"]);
                    //    break;
                    case "dbclick":
                        action_present = true;
                        jsonTree1.Add(id.ToString(), jsonTree2);

                        jsonTree2.Add("id", id);
                        jsonTree2.Add("type", "CLICK_EL");
                        jsonTree2.Add("parentId", parent_id);
                        jsonTree2.Add("label", "Double Click");
                        jsonTree2.Add("key", "dclick");
                        jsonTree2.Add("allowNesting", false);
                        jsonTree2.Add("selector", "");
                        jsonTree2.Add("selectorType", "location");
                        jsonTree2.Add("elSelector", "");
                        jsonTree2.Add("imgSelector", $"{item["image"]}");
                        jsonTree2.Add("instruction", $"vision MouseClick, left, {item["coord"]["x"]}, {item["coord"]["y"]}");
                        jsonTree2.Add("description", $"Click left, {item["coord"]["x"]}, {item["coord"]["y"]}");
                        jsonTree2.Add("status", true);
                        jsonTree2.Add("breakpoint", false);
                        jsonTree2.Add("x", item["coord"]["x"]);
                        jsonTree2.Add("y", item["coord"]["y"]);
                        break;
                    case "open":
                        action_present = true;
                        jsonTree1.Add(id.ToString(), jsonTree2);

                        jsonTree2.Add("id", id);
                        jsonTree2.Add("type", "OPEN_WEB_PAGE");
                        jsonTree2.Add("parentId", parent_id);
                        jsonTree2.Add("label", "Open Webpage");
                        jsonTree2.Add("key", "openwebpage");
                        jsonTree2.Add("allowNesting", false);
                        jsonTree2.Add("subActions", jsonActionArrayString);
                        jsonTree2.Add("url", item["value"]);
                        jsonTree2.Add("path", "");
                        jsonTree2.Add("browser", "chrome");
                        jsonTree2.Add("instruction", item["value"]);
                        jsonTree2.Add("description", item["value"]);
                        jsonTree2.Add("status", true);
                        jsonTree2.Add("breakpoint", false);
                        break;
                    case "type":
                    case "kbd":
                        action_present = true;
                        jsonTree1.Add(id.ToString(), jsonTree2);

                        jsonTree2.Add("id", id);
                        jsonTree2.Add("type", "TYPE_TEXT_EL");
                        jsonTree2.Add("parentId", parent_id);
                        jsonTree2.Add("label", "Enter Keystrokes");
                        jsonTree2.Add("key", "keyboard");
                        jsonTree2.Add("selectorType", "elSelector");
                        jsonTree2.Add("resultType", "string");
                        jsonTree2.Add("xpath", item["xPath"]);
                        jsonTree2.Add("text", item["value"]);
                        jsonTree2.Add("imgSelector", null);
                        jsonTree2.Add("imgSVG", item["img"]);
                        jsonTree2.Add("allowNesting", false);
                        jsonTree2.Add("instruction", $"type {item["xPath"]} as {item["value"]}");
                        jsonTree2.Add("description", $"{item["value"]} in {item["xPath"]}");
                        jsonTree2.Add("status", true);
                        jsonTree2.Add("breakpoint", false);
                        jsonTree2.Add("inputPlaceHolder", $"{item["placeholder"]}");
                        jsonTree2.Add("inputLabel", $"{item["label"]}");
                        jsonTree2.Add("inputName", $"{item["name"]}");
                        break;
                    case "kbd_shortcut":
                        bool ctrl = false, shift = false, alt = false, win = false;
                        string desc = "";
                        List<string> vals = new List<string>();
                        foreach (var jtok in item["value"])
                        {
                            if (jtok.ToString() == "ctrl") ctrl = true;
                            else if (jtok.ToString() == "shift") shift = true;
                            else if (jtok.ToString() == "alt") alt = true;
                            else if (jtok.ToString() == "win") win = true;
                            else vals.Add(jtok.ToString());
                            desc += jtok.ToString() + " + ";
                        }
                        char[] charstoTrim = { ' ', '+' };
                        desc.TrimEnd(charstoTrim);
                        action_present = true;
                        jsonTree1.Add(id.ToString(), jsonTree2);

                        jsonTree2.Add("id", id);
                        jsonTree2.Add("type", "KEY_STROKES");
                        jsonTree2.Add("parentId", parent_id);
                        jsonTree2.Add("label", "Keyboard ShortCuts");
                        jsonTree2.Add("key", "keyboardshortcuts");
                        jsonTree2.Add("win", win);
                        jsonTree2.Add("alt", alt);
                        jsonTree2.Add("ctrl", ctrl);
                        jsonTree2.Add("shift", shift);
                        jsonTree2.Add("selector", "custom");
                        jsonTree2.Add("keyValue", "");
                        jsonTree2.Add("value", item["value"].Last);
                        jsonTree2.Add("instruction", $"vision Send, #{item["value"].Last}");
                        jsonTree2.Add("description", desc);
                        jsonTree2.Add("status", true);
                        jsonTree2.Add("breakpoint", false);
                        break;
                    case "kp":
                    case "kh":
                        action_present = true;
                        jsonTree1.Add(id.ToString(), jsonTree2);

                        jsonTree2.Add("id", id);
                        jsonTree2.Add("type", "KEY_STROKES");
                        jsonTree2.Add("parentId", parent_id);
                        jsonTree2.Add("label", "Keyboard ShortCuts");
                        jsonTree2.Add("key", "keyboardshortcuts");
                        jsonTree2.Add("win", false);
                        jsonTree2.Add("alt", false);
                        jsonTree2.Add("ctrl", false);
                        jsonTree2.Add("shift", false);
                        jsonTree2.Add("selector", "custom");
                        jsonTree2.Add("keyValue", "");
                        jsonTree2.Add("value", item["value"]);
                        jsonTree2.Add("instruction", $"vision Send, #{item["value"]}");
                        jsonTree2.Add("description", item["value"]);
                        jsonTree2.Add("status", true);
                        jsonTree2.Add("breakpoint", false);
                        break;
                }

                if (action_present) jsonArrayString.Add(id.ToString());
                action_present = false;
            }

            var tobj = new object();
            var tempList = new List<string>();
            jsonTree0.Add("variables", tobj);
            jsonTree0.Add("parameters", tobj);
            jsonTree0.Add("ftpConfig", tobj);
            jsonTree0.Add("connect_tree", tobj);
            tempList.Add(parent_id.ToString());
            jsonTree0.Add("sequence", tempList);
            tempList = new List<string>
            {
                jsonArrayString[0]
            };
            jsonTree0.Add("selectedActions", tempList);
            jsonTree0.Add("flowPath", "");
            jsonTree0.Add("dirPath", "");
            jsonTree0.Add("logs", "");
            jsonTree0.Add("currentAction", jsonArrayString[0]);
            jsonTree0.Add("isEdit", true);*/

            var resultant = JsonConvert.SerializeObject(masterObject, Formatting.Indented);

            return new ConversionResult { Status = true, Data = resultant };
        }
    }
    public struct ConversionResult
    {
        public bool Status { get; set; }
        public string Data { get; set; }
    }
}
