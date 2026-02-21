using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using ReleasePack.Engine;

namespace ReleasePack.AddIn.McpServer
{
    /// <summary>
    /// Embedded MCP (Model Context Protocol) server running inside SolidWorks.
    /// Listens on http://localhost:8123/mcp for JSON-RPC tool calls from AI agents.
    /// All SolidWorks COM calls are marshaled to the main STA thread via SynchronizationContext.
    /// </summary>
    public class SwMcpServer : IDisposable
    {
        private readonly ISldWorks _swApp;
        private readonly SynchronizationContext _uiContext;
        private HttpListener _listener;
        private CancellationTokenSource _cts;
        private Thread _serverThread;
        private readonly JavaScriptSerializer _json = new JavaScriptSerializer();

        public const int DEFAULT_PORT = 8123;
        public const string PREFIX = "http://localhost:{0}/mcp/";

        public bool IsRunning { get; private set; }

        // Available tools exposed to AI agents
        private readonly Dictionary<string, Func<Dictionary<string, object>, object>> _tools;

        public SwMcpServer(ISldWorks swApp)
        {
            _swApp = swApp;
            // Capture the UI thread's SynchronizationContext so we can marshal COM calls
            _uiContext = SynchronizationContext.Current
                ?? new System.Windows.Forms.WindowsFormsSynchronizationContext();

            _tools = new Dictionary<string, Func<Dictionary<string, object>, object>>(StringComparer.OrdinalIgnoreCase)
            {
                { "get_active_model_info", Tool_GetActiveModelInfo },
                { "list_open_documents", Tool_ListOpenDocuments },
                { "select_entity", Tool_SelectEntity },
                { "clear_selection", Tool_ClearSelection },
                { "generate_drawing", Tool_GenerateDrawing },
                { "export_active_document", Tool_ExportActiveDocument },
                { "get_mass_properties", Tool_GetMassProperties },
                { "get_bounding_box", Tool_GetBoundingBox },
                { "get_custom_properties", Tool_GetCustomProperties },
                { "set_custom_property", Tool_SetCustomProperty },
                { "save_document", Tool_SaveDocument },
                { "run_macro", Tool_RunMacro },
                { "get_component_tree", Tool_GetComponentTree },
                { "get_feature_list", Tool_GetFeatureList },
            };
        }

        #region Server Lifecycle

        /// <summary>Start the MCP server on the specified port.</summary>
        public void Start(int port = DEFAULT_PORT)
        {
            if (IsRunning) return;

            _cts = new CancellationTokenSource();
            string prefix = string.Format(PREFIX, port);

            _serverThread = new Thread(() => RunServer(prefix, _cts.Token))
            {
                IsBackground = true,
                Name = "SwMcpServer"
            };
            _serverThread.Start();
            IsRunning = true;

            System.Diagnostics.Trace.WriteLine($"[MCP] Server started on {prefix}");
        }

        /// <summary>Stop the MCP server gracefully.</summary>
        public void Stop()
        {
            if (!IsRunning) return;

            _cts?.Cancel();
            _listener?.Stop();
            _listener?.Close();
            IsRunning = false;

            System.Diagnostics.Trace.WriteLine("[MCP] Server stopped.");
        }

        public void Dispose()
        {
            Stop();
            _cts?.Dispose();
        }

        #endregion

        #region HTTP Loop

        private void RunServer(string prefix, CancellationToken ct)
        {
            try
            {
                _listener = new HttpListener();
                _listener.Prefixes.Add(prefix);
                _listener.Start();

                while (!ct.IsCancellationRequested)
                {
                    try
                    {
                        // GetContext blocks until a request arrives
                        var context = _listener.GetContext();
                        // Handle each request on a thread pool thread
                        ThreadPool.QueueUserWorkItem(_ => HandleRequest(context));
                    }
                    catch (HttpListenerException) when (ct.IsCancellationRequested)
                    {
                        break; // Listener was stopped
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Trace.WriteLine($"[MCP] Error: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine($"[MCP] Fatal error: {ex.Message}");
            }
        }

        private void HandleRequest(HttpListenerContext context)
        {
            string responseBody;
            int statusCode = 200;

            try
            {
                // CORS headers for browser-based clients
                context.Response.Headers.Add("Access-Control-Allow-Origin", "*");
                context.Response.Headers.Add("Access-Control-Allow-Methods", "POST, GET, OPTIONS");
                context.Response.Headers.Add("Access-Control-Allow-Headers", "Content-Type");

                if (context.Request.HttpMethod == "OPTIONS")
                {
                    context.Response.StatusCode = 204;
                    context.Response.Close();
                    return;
                }

                string path = context.Request.Url.AbsolutePath.TrimEnd('/');

                if (context.Request.HttpMethod == "GET" && path == "/mcp")
                {
                    // Health check / tool list
                    responseBody = _json.Serialize(new
                    {
                        server = "SolidWorks MCP Server",
                        version = "1.0.0",
                        status = "running",
                        solidworks_connected = _swApp != null,
                        available_tools = new List<string>(_tools.Keys)
                    });
                }
                else if (context.Request.HttpMethod == "POST" && path == "/mcp/tools/call")
                {
                    // JSON-RPC style tool call
                    string body;
                    using (var reader = new StreamReader(context.Request.InputStream, context.Request.ContentEncoding))
                    {
                        body = reader.ReadToEnd();
                    }

                    var request = _json.Deserialize<Dictionary<string, object>>(body);
                    responseBody = ExecuteToolCall(request);
                }
                else if (context.Request.HttpMethod == "GET" && path == "/mcp/tools/list")
                {
                    // Return tool descriptions
                    responseBody = _json.Serialize(GetToolDescriptions());
                }
                else
                {
                    statusCode = 404;
                    responseBody = _json.Serialize(new { error = "Not found", path });
                }
            }
            catch (Exception ex)
            {
                statusCode = 500;
                responseBody = _json.Serialize(new { error = ex.Message, type = ex.GetType().Name });
            }

            try
            {
                byte[] buffer = Encoding.UTF8.GetBytes(responseBody);
                context.Response.StatusCode = statusCode;
                context.Response.ContentType = "application/json";
                context.Response.ContentLength64 = buffer.Length;
                context.Response.OutputStream.Write(buffer, 0, buffer.Length);
                context.Response.Close();
            }
            catch { }
        }

        #endregion

        #region Tool Execution (Marshaled to STA Thread)

        private string ExecuteToolCall(Dictionary<string, object> request)
        {
            if (request == null || !request.ContainsKey("tool"))
                return _json.Serialize(new { error = "Missing 'tool' field in request body" });

            string toolName = request["tool"].ToString();

            if (!_tools.ContainsKey(toolName))
                return _json.Serialize(new { error = $"Unknown tool: {toolName}", available = new List<string>(_tools.Keys) });

            var args = request.ContainsKey("arguments")
                ? request["arguments"] as Dictionary<string, object> ?? new Dictionary<string, object>()
                : new Dictionary<string, object>();

            // Marshal the tool call to the main STA thread where COM objects live
            object result = null;
            Exception error = null;
            var done = new ManualResetEventSlim(false);

            _uiContext.Post(_ =>
            {
                try
                {
                    result = _tools[toolName](args);
                }
                catch (Exception ex)
                {
                    error = ex;
                }
                finally
                {
                    done.Set();
                }
            }, null);

            // Wait up to 30 seconds for the tool to complete
            if (!done.Wait(TimeSpan.FromSeconds(30)))
            {
                return _json.Serialize(new { error = "Tool call timed out after 30 seconds", tool = toolName });
            }

            if (error != null)
            {
                return _json.Serialize(new { error = error.Message, tool = toolName, type = error.GetType().Name });
            }

            return _json.Serialize(new { tool = toolName, result });
        }

        #endregion

        #region Tool Implementations

        private object Tool_GetActiveModelInfo(Dictionary<string, object> args)
        {
            ModelDoc2 doc = (ModelDoc2)_swApp.ActiveDoc;
            if (doc == null) return new { error = "No active document" };

            int docType = doc.GetType();
            string typeName = docType == (int)swDocumentTypes_e.swDocPART ? "Part" :
                              docType == (int)swDocumentTypes_e.swDocASSEMBLY ? "Assembly" :
                              docType == (int)swDocumentTypes_e.swDocDRAWING ? "Drawing" : "Unknown";

            return new
            {
                title = doc.GetTitle(),
                path = doc.GetPathName(),
                type = typeName,
                configuration = ((Configuration)doc.GetActiveConfiguration())?.Name,
                saved = !doc.GetSaveFlag(),
            };
        }

        private object Tool_ListOpenDocuments(Dictionary<string, object> args)
        {
            var docs = new List<object>();
            ModelDoc2 doc = (ModelDoc2)_swApp.GetFirstDocument();
            while (doc != null)
            {
                docs.Add(new { title = doc.GetTitle(), path = doc.GetPathName() });
                doc = (ModelDoc2)doc.GetNext();
            }
            return docs;
        }

        private object Tool_SelectEntity(Dictionary<string, object> args)
        {
            ModelDoc2 doc = (ModelDoc2)_swApp.ActiveDoc;
            if (doc == null) return new { error = "No active document" };

            string name = args.ContainsKey("name") ? args["name"].ToString() : null;
            string type = args.ContainsKey("type") ? args["type"].ToString() : null;
            bool append = args.ContainsKey("append") && Convert.ToBoolean(args["append"]);

            if (string.IsNullOrEmpty(name))
                return new { error = "Missing 'name' argument" };

            bool success = doc.Extension.SelectByID2(
                name, type ?? "", 0, 0, 0, append, 0, null, 0);

            return new { selected = success, entity = name };
        }

        private object Tool_ClearSelection(Dictionary<string, object> args)
        {
            ModelDoc2 doc = (ModelDoc2)_swApp.ActiveDoc;
            if (doc == null) return new { error = "No active document" };

            doc.ClearSelection2(true);
            return new { cleared = true };
        }

        private object Tool_GenerateDrawing(Dictionary<string, object> args)
        {
            ModelDoc2 doc = (ModelDoc2)_swApp.ActiveDoc;
            if (doc == null) return new { error = "No active document" };

            var options = new ExportOptions
            {
                GenerateDrawing = true,
                ExportPDF = args.ContainsKey("pdf") && Convert.ToBoolean(args["pdf"]),
                ExportDXF = args.ContainsKey("dxf") && Convert.ToBoolean(args["dxf"]),
            };

            if (args.ContainsKey("output_folder"))
            {
                options.OutputFolder = args["output_folder"].ToString();
                options.UseCustomFolder = true;
            }

            if (args.ContainsKey("dimension_mode"))
            {
                string mode = args["dimension_mode"].ToString().ToLower();
                options.DimensionMode = mode == "model" ? DimensionMode.ModelDimensions :
                                        mode == "hybrid" ? DimensionMode.HybridAuto :
                                        DimensionMode.FullAuto;
            }

            var pipeline = new ExportPipeline(_swApp);
            var results = pipeline.Execute(options);

            return new
            {
                success = true,
                generated = results?.Count ?? 0,
                files = results
            };
        }

        private object Tool_ExportActiveDocument(Dictionary<string, object> args)
        {
            ModelDoc2 doc = (ModelDoc2)_swApp.ActiveDoc;
            if (doc == null) return new { error = "No active document" };

            string format = args.ContainsKey("format") ? args["format"].ToString().ToUpper() : "STEP";
            string outputPath = args.ContainsKey("output_path") ? args["output_path"].ToString() : null;

            if (string.IsNullOrEmpty(outputPath))
            {
                string basePath = Path.GetDirectoryName(doc.GetPathName());
                string baseName = Path.GetFileNameWithoutExtension(doc.GetPathName());
                string ext = format == "PDF" ? ".pdf" : format == "DXF" ? ".dxf" :
                             format == "PARASOLID" ? ".x_t" : ".step";
                outputPath = Path.Combine(basePath, baseName + ext);
            }

            int errors = 0, warnings = 0;
            bool success = doc.Extension.SaveAs2(
                outputPath, 0, 0, null, "", false, ref errors, ref warnings);

            return new { success, output_path = outputPath, errors, warnings };
        }

        private object Tool_GetMassProperties(Dictionary<string, object> args)
        {
            ModelDoc2 doc = (ModelDoc2)_swApp.ActiveDoc;
            if (doc == null) return new { error = "No active document" };

            var ext = (ModelDocExtension)doc.Extension;
            int status;
            double[] props = (double[])ext.GetMassProperties2(1, out status, false);

            if (props == null || props.Length < 12)
                return new { error = "Could not retrieve mass properties", status };

            return new
            {
                center_of_mass = new[] { props[0], props[1], props[2] },
                volume_m3 = props[3],
                surface_area_m2 = props[4],
                mass_kg = props[5],
                moments_of_inertia = new[] { props[6], props[7], props[8], props[9], props[10], props[11] }
            };
        }

        private object Tool_GetBoundingBox(Dictionary<string, object> args)
        {
            ModelDoc2 doc = (ModelDoc2)_swApp.ActiveDoc;
            if (doc == null) return new { error = "No active document" };

            double[] bbox = null;
            try
            {
                if (doc.GetType() == (int)swDocumentTypes_e.swDocPART)
                {
                    PartDoc part = (PartDoc)doc;
                    object[] bodies = (object[])part.GetBodies2((int)swBodyType_e.swSolidBody, true);
                    if (bodies != null && bodies.Length > 0)
                    {
                        IBody2 body = (IBody2)bodies[0];
                        bbox = (double[])body.GetBodyBox();
                    }
                }
            }
            catch { }

            if (bbox == null || bbox.Length < 6)
                return new { error = "Could not retrieve bounding box - only supported for active Part documents with solid bodies" };

            return new
            {
                min = new[] { bbox[0], bbox[1], bbox[2] },
                max = new[] { bbox[3], bbox[4], bbox[5] },
                size_m = new[] {
                    Math.Abs(bbox[3] - bbox[0]),
                    Math.Abs(bbox[4] - bbox[1]),
                    Math.Abs(bbox[5] - bbox[2])
                }
            };
        }

        private object Tool_GetCustomProperties(Dictionary<string, object> args)
        {
            ModelDoc2 doc = (ModelDoc2)_swApp.ActiveDoc;
            if (doc == null) return new { error = "No active document" };

            var ext = (ModelDocExtension)doc.Extension;
            var mgr = ext.CustomPropertyManager[""];

            object names = null, types = null, values = null, resolved = null;
            mgr.GetAll2(ref names, ref types, ref values, ref resolved);

            var propNames = names as string[];
            var propVals = resolved as string[];

            var result = new Dictionary<string, string>();
            if (propNames != null)
            {
                for (int i = 0; i < propNames.Length; i++)
                    result[propNames[i]] = propVals != null && i < propVals.Length ? propVals[i] : "";
            }
            return result;
        }

        private object Tool_SetCustomProperty(Dictionary<string, object> args)
        {
            ModelDoc2 doc = (ModelDoc2)_swApp.ActiveDoc;
            if (doc == null) return new { error = "No active document" };

            string name = args.ContainsKey("name") ? args["name"].ToString() : null;
            string value = args.ContainsKey("value") ? args["value"].ToString() : "";

            if (string.IsNullOrEmpty(name))
                return new { error = "Missing 'name' argument" };

            var mgr = ((ModelDocExtension)doc.Extension).CustomPropertyManager[""];
            int ret = mgr.Add3(name, (int)swCustomInfoType_e.swCustomInfoText, value,
                (int)swCustomPropertyAddOption_e.swCustomPropertyReplaceValue);

            return new { success = ret == 0 || ret == 1, property = name, value };
        }

        private object Tool_SaveDocument(Dictionary<string, object> args)
        {
            ModelDoc2 doc = (ModelDoc2)_swApp.ActiveDoc;
            if (doc == null) return new { error = "No active document" };

            int errors = 0, warnings = 0;
            bool success = doc.Save3(
                (int)swSaveAsOptions_e.swSaveAsOptions_Silent, ref errors, ref warnings);

            return new { success, errors, warnings };
        }

        private object Tool_RunMacro(Dictionary<string, object> args)
        {
            string macroPath = args.ContainsKey("path") ? args["path"].ToString() : null;
            string module = args.ContainsKey("module") ? args["module"].ToString() : "";
            string procedure = args.ContainsKey("procedure") ? args["procedure"].ToString() : "main";

            if (string.IsNullOrEmpty(macroPath))
                return new { error = "Missing 'path' argument" };

            if (!File.Exists(macroPath))
                return new { error = $"Macro file not found: {macroPath}" };

            int err;
            bool success = _swApp.RunMacro2(macroPath, module, procedure, 
                (int)swRunMacroOption_e.swRunMacroDefault, out err);

            return new { success, error_code = err };
        }

        private object Tool_GetComponentTree(Dictionary<string, object> args)
        {
            ModelDoc2 doc = (ModelDoc2)_swApp.ActiveDoc;
            if (doc == null) return new { error = "No active document" };

            if (doc.GetType() != (int)swDocumentTypes_e.swDocASSEMBLY)
                return new { error = "Active document is not an assembly" };

            AssemblyDoc assy = (AssemblyDoc)doc;
            object[] components = (object[])assy.GetComponents(true); // Top-level only

            var tree = new List<object>();
            if (components != null)
            {
                foreach (Component2 comp in components)
                {
                    try
                    {
                        tree.Add(new
                        {
                            name = comp.Name2,
                            path = comp.GetPathName(),
                            suppressed = comp.GetSuppression2() == (int)swComponentSuppressionState_e.swComponentSuppressed,
                            visible = comp.Visible == (int)swComponentVisibilityState_e.swComponentVisible
                        });
                    }
                    catch { }
                }
            }
            return tree;
        }

        private object Tool_GetFeatureList(Dictionary<string, object> args)
        {
            ModelDoc2 doc = (ModelDoc2)_swApp.ActiveDoc;
            if (doc == null) return new { error = "No active document" };

            var features = new List<object>();
            Feature feat = (Feature)doc.FirstFeature();
            while (feat != null)
            {
                try
                {
                    features.Add(new
                    {
                        name = feat.Name,
                        type = feat.GetTypeName2(),
                        suppressed = ((bool[])feat.IsSuppressed2(
                            (int)swInConfigurationOpts_e.swThisConfiguration, null))[0]
                    });
                }
                catch { }
                feat = (Feature)feat.GetNextFeature();
            }
            return features;
        }

        #endregion

        #region Tool Descriptions

        private object GetToolDescriptions()
        {
            return new[]
            {
                new { name = "get_active_model_info", description = "Get info about the active SolidWorks document (title, path, type, configuration)", arguments = new string[0] },
                new { name = "list_open_documents", description = "List all currently open documents in SolidWorks", arguments = new string[0] },
                new { name = "select_entity", description = "Select an entity (face, edge, feature, component) by name", arguments = new[] { "name (required)", "type (optional)", "append (optional, bool)" } },
                new { name = "clear_selection", description = "Clear the current selection in the active document", arguments = new string[0] },
                new { name = "generate_drawing", description = "Generate a 2D drawing from the active model using Release Pack engine", arguments = new[] { "pdf (optional, bool)", "dxf (optional, bool)", "output_folder (optional)", "dimension_mode (optional: auto|model|hybrid)" } },
                new { name = "export_active_document", description = "Export the active document to a specified format", arguments = new[] { "format (STEP|PDF|DXF|PARASOLID)", "output_path (optional)" } },
                new { name = "get_mass_properties", description = "Get mass, volume, surface area, center of mass, and moments of inertia", arguments = new string[0] },
                new { name = "get_bounding_box", description = "Get the bounding box (min/max coordinates) of the active model", arguments = new string[0] },
                new { name = "get_custom_properties", description = "Get all custom properties of the active document", arguments = new string[0] },
                new { name = "set_custom_property", description = "Set a custom property value on the active document", arguments = new[] { "name (required)", "value (required)" } },
                new { name = "save_document", description = "Save the active document", arguments = new string[0] },
                new { name = "run_macro", description = "Execute a SolidWorks macro (.swp) file", arguments = new[] { "path (required)", "module (optional)", "procedure (optional, default: main)" } },
                new { name = "get_component_tree", description = "Get the top-level component tree of the active assembly", arguments = new string[0] },
                new { name = "get_feature_list", description = "List all features in the active model's feature tree", arguments = new string[0] },
            };
        }

        #endregion
    }
}
