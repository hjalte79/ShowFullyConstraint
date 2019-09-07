////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Copyright (c) Hjälte. All rights reserved 
// Written by Jelte de Jong - 2019
//
// Permission to use, copy, modify, and distribute this software in
// object code form for any purpose and without fee is hereby granted, 
// provided that the above copyright notice appears in all copies and 
// that both that copyright notice and the limited warranty and
// restricted rights notice below appear in all supporting 
// documentation.
//
// HJALTE PROVIDES THIS PROGRAM "AS IS" AND WITH ALL FAULTS. 
// HJALTE SPECIFICALLY DISCLAIMS ANY IMPLIED WARRANTY OF
// MERCHANTABILITY OR FITNESS FOR A PARTICULAR USE.  
// DOES NOT WARRANT THAT THE OPERATION OF THE PROGRAM WILL BE
// UNINTERRUPTED OR ERROR FREE.
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Inventor;

namespace Hjalte.ShowFullyConstraint
{
    [GuidAttribute("d53b3492-2496-4816-b14d-903992a52a5e"), ComVisible(true)]
    public class StandardAddInServer : Inventor.ApplicationAddInServer
    {
        //https://forums.autodesk.com/t5/inventor-ideas/change-icon-of-fully-constraint-parts-in-assembly/idi-p/4600123

        private Inventor.Application inventor;

        public void Activate(Inventor.ApplicationAddInSite addInSiteObject, bool firstTime)
        {
            inventor = addInSiteObject.Application;


            inventor.ApplicationEvents.OnDocumentChange += onChange;
            inventor.ApplicationEvents.OnSaveDocument += onSaving;
            inventor.ApplicationEvents.OnCloseDocument += onClosing;
            inventor.ApplicationEvents.OnOpenDocument += onOpen;
        }
        private bool isSaving = false;

        private void onChange(_Document DocumentObject, EventTimingEnum BeforeOrAfter, CommandTypesEnum ReasonsForChange, NameValueMap Context, out HandlingCodeEnum HandlingCode)
        {
            if (BeforeOrAfter == EventTimingEnum.kAfter)
            {
                startMarking(DocumentObject);
            }
            HandlingCode = HandlingCodeEnum.kEventNotHandled;
        }
        private void onOpen(_Document DocumentObject, string FullDocumentName, EventTimingEnum BeforeOrAfter, NameValueMap Context, out HandlingCodeEnum HandlingCode)
        {
            if (BeforeOrAfter == EventTimingEnum.kAfter)
            {
                // startMarking(DocumentObject);
            }
            HandlingCode = HandlingCodeEnum.kEventNotHandled;
        }
        private void onClosing(_Document DocumentObject, string FullDocumentName, EventTimingEnum BeforeOrAfter, NameValueMap Context, out HandlingCodeEnum HandlingCode)
        {
            if (BeforeOrAfter == EventTimingEnum.kBefore)
            {
                startUnMarking(DocumentObject);
            }
            HandlingCode = HandlingCodeEnum.kEventNotHandled;
        }
        private void onSaving(_Document DocumentObject, EventTimingEnum BeforeOrAfter, NameValueMap Context, out HandlingCodeEnum HandlingCode)
        {
            if (BeforeOrAfter == EventTimingEnum.kBefore)
            {
                isSaving = true;
                startUnMarking(DocumentObject);
            }
            else
            {
                isSaving = false;
            }
            HandlingCode = HandlingCodeEnum.kEventNotHandled;
        }

        
        private void startUnMarking(Document doc)
        {
            if (doc.DocumentType != DocumentTypeEnum.kAssemblyDocumentObject)
            {
                return;
            }
            AssemblyDocument aDoc = (AssemblyDocument)doc;
            unMark(aDoc.BrowserPanes["Model"].TopNode.BrowserNodes);
        }
        private void unMark(BrowserNodesEnumerator browserNodes)
        {
            foreach (BrowserNode node in browserNodes)
            {
                BrowserNodeDisplayStateEnum displaystate = node.BrowserNodeDefinition.DisplayState;
                if (displaystate == BrowserNodeDisplayStateEnum.kDefaultDisplayState ||
                        displaystate == BrowserNodeDisplayStateEnum.kGreenCheckDisplayState ||
                        displaystate == BrowserNodeDisplayStateEnum.kCyclicDisplayState)
                {
                    node.BrowserNodeDefinition.DisplayState = BrowserNodeDisplayStateEnum.kDefaultDisplayState;
                }

                if (node.NativeObject is BrowserFolder)
                {
                    unMark(node.BrowserNodes);
                }
            }

        }

        private void startMarking(Document doc)
        {
            if (isSaving)
            {
                return;
            }
            if (doc is null)
            {
                return;
            }
            if (doc.DocumentType != DocumentTypeEnum.kAssemblyDocumentObject)
            {
                return;
            }
            AssemblyDocument aDoc = (AssemblyDocument)doc;

            mark(aDoc.BrowserPanes["Model"].TopNode.BrowserNodes);

        }
        private void mark(BrowserNodesEnumerator browserNodes)
        {
            foreach (BrowserNode node in browserNodes)
            {
                if (node.NativeObject is BrowserFolder)
                {
                    mark(node.BrowserNodes);
                }

                if (node.NativeObject is ComponentOccurrence)
                {
                    ComponentOccurrence occ = (ComponentOccurrence)node.NativeObject;
                    int TranslationDegreesCount;
                    ObjectsEnumerator TranslationDegreesVectors;
                    int RotationDegreesCount;
                    ObjectsEnumerator RotationDegreesVectors;
                    Point DOFCenter;

                    occ.GetDegreesOfFreedom(out TranslationDegreesCount, out TranslationDegreesVectors,
                                        out RotationDegreesCount, out RotationDegreesVectors, out DOFCenter);

                    BrowserNodeDisplayStateEnum displaystate = node.BrowserNodeDefinition.DisplayState;
                    if (displaystate == BrowserNodeDisplayStateEnum.kDefaultDisplayState || 
                        displaystate == BrowserNodeDisplayStateEnum.kGreenCheckDisplayState ||
                        displaystate == BrowserNodeDisplayStateEnum.kCyclicDisplayState)
                    {
                        if ((TranslationDegreesCount == 0) && (RotationDegreesCount == 0))
                        {
                            node.BrowserNodeDefinition.DisplayState = BrowserNodeDisplayStateEnum.kGreenCheckDisplayState;
                        }
                        //else if ((TranslationDegreesCount == 0) && (RotationDegreesCount == 1))
                        //{
                        //    node.BrowserNodeDefinition.DisplayState = BrowserNodeDisplayStateEnum.kCyclicDisplayState;
                        //}
                        else
                        {
                            node.BrowserNodeDefinition.DisplayState = BrowserNodeDisplayStateEnum.kDefaultDisplayState;
                        }
                    }
                }
            }
        }

        public void Deactivate()
        {
            // Release objects.
            Marshal.ReleaseComObject(inventor);
            inventor = null;

            GC.WaitForPendingFinalizers();
            GC.Collect();
        }
        public void ExecuteCommand(int commandID)
        {
            // Note:this method is now obsolete, you should use the 
            // ControlDefinition functionality for implementing commands.
        }
        public object Automation
        {
            get
            {
                // TODO: Add ApplicationAddInServer.Automation getter implementation
                return null;
            }
        }
    }
}
