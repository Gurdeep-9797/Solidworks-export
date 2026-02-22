using System;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace ReleasePack.Engine.Exporting
{
    /// <summary>
    /// V2 Large Assembly Memory Controller.
    /// Handles invisible document opening, graphic data suppression, and aggressive .NET/COM Garbage Collection
    /// required to process 5000+ part assemblies over 8+ hour workloads without crashing SolidWorks.
    /// </summary>
    public class MemoryController
    {
        private readonly ISldWorks _swApp;
        private int _totalDocsProcessed = 0;

        public MemoryController(ISldWorks swApp)
        {
            _swApp = swApp;
        }

        /// <summary>
        /// Attempts to open a document invisibly to prevent GDI+ handle exhaustion and memory spikes.
        /// </summary>
        public ModelDoc2 OpenDocumentInvisible(string filePath)
        {
            if (!System.IO.File.Exists(filePath)) return null;

            int docType = GetSwDocumentType(filePath);
            if (docType == (int)swDocumentTypes_e.swDocNONE) return null;

            // Start with base document open specification
            IDocumentSpecification docSpec = (IDocumentSpecification)_swApp.GetOpenDocSpec(filePath);
            
            // Critical Memory Flags
            docSpec.Silent = true;       // No popup dialogs
            docSpec.ReadOnly = true;     // We only read geometry to generate drawings
            
            // To prevent UI resource starvation, keep the document hidden
            // WARNING: Opening drawing files entirely hidden can sometimes fail bounding box calculations
            // But for SLDPRT it is 100% vital
            _swApp.DocumentVisible(false, docType); 

            ModelDoc2 openedDoc = _swApp.OpenDoc7(docSpec);

            // Re-enable document visibility generically so user-clicked models still open correctly later
            _swApp.DocumentVisible(true, (int)swDocumentTypes_e.swDocPART);
            _swApp.DocumentVisible(true, (int)swDocumentTypes_e.swDocASSEMBLY);
            _swApp.DocumentVisible(true, (int)swDocumentTypes_e.swDocDRAWING);

            return openedDoc;
        }

        /// <summary>
        /// Cleans up pointer locks and forces garbage collection. 
        /// Crucial when iterating an assembly tree.
        /// </summary>
        public void ReleaseDocument(ModelDoc2 doc)
        {
            if (doc != null)
            {
                string title = doc.GetTitle();
                _swApp.CloseDoc(title);
            }

            _totalDocsProcessed++;

            // SolidWorks leaks memory heavily through COM references.
            // Aggressive sweep every 50 documents processed.
            if (_totalDocsProcessed % 50 == 0)
            {
                ForceGarbageCollection();
            }
        }

        /// <summary>
        /// Extreme multi-generation sweep for releasing orphaned COM objects.
        /// </summary>
        public static void ForceGarbageCollection()
        {
            GC.Collect(GC.MaxGeneration, GCCollectionMode.Forced, true);
            GC.WaitForPendingFinalizers();

            // Run twice for stubborn unmanaged RCW wrappers
            GC.Collect(GC.MaxGeneration, GCCollectionMode.Forced, true);
            GC.WaitForPendingFinalizers();
        }

        private int GetSwDocumentType(string filePath)
        {
            string ext = System.IO.Path.GetExtension(filePath).ToLower();
            switch (ext)
            {
                case ".sldprt": return (int)swDocumentTypes_e.swDocPART;
                case ".sldasm": return (int)swDocumentTypes_e.swDocASSEMBLY;
                case ".slddrw": return (int)swDocumentTypes_e.swDocDRAWING;
                default: return (int)swDocumentTypes_e.swDocNONE;
            }
        }
    }
}
