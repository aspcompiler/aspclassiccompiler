// ----------------------------------------------------------------------------------
// Microsoft Developer & Platform Evangelism
// 
// Copyright (c) Microsoft Corporation. All rights reserved.
// 
// THIS CODE AND INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, 
// EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES 
// OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
// ----------------------------------------------------------------------------------
// The example companies, organizations, products, domain names,
// e-mail addresses, logos, people, places, and events depicted
// herein are fictitious.  No association with any real company,
// organization, product, domain name, email address, logo, person,
// places, or events is intended or should be inferred.
// ----------------------------------------------------------------------------------

//
// <copyright file="BlobProvider.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
//
using System;
using System.Collections.Generic;
using System.IO;
using System.Configuration;
using Microsoft.ServiceHosting.ServiceRuntime;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Globalization;
using Microsoft.Samples.ServiceHosting.StorageClient;


namespace Microsoft.Samples.ServiceHosting.AspProviders
{

     public enum EventKind
     {
          Critical,
          Error,
          Warning,
          Information,
          Verbose
     };

    static class Log
    {
        internal static void WriteToLog(string logName, string fmt, params object[] args)
        {
            RoleManager.WriteToLog(logName, String.Format(CultureInfo.InvariantCulture, fmt, args));
        }
        internal static void Write(EventKind eventKind, string message, params object[] args)
        {
            if (RoleManager.IsRoleManagerRunning)
            {
                switch(eventKind)
                {
                    case EventKind.Error:
                        WriteToLog("Error", message, args);
                        break;
                    case EventKind.Critical:
                        WriteToLog("Critical", message, args);
                        break;
                    case EventKind.Warning:
                        WriteToLog("Warning", message, args);
                        break;
                    case EventKind.Information:
                        WriteToLog("Information", message, args);
                        break;
                    case EventKind.Verbose:
                        WriteToLog("Verbose", message, args);
                        break;
                }
            }
            else
            {
                switch (eventKind)
                {
                    case EventKind.Error:
                    case EventKind.Critical:
                        Trace.TraceError(message, args);
                        break;
                    case EventKind.Warning:
                        Trace.TraceWarning(message, args);
                        break;
                    case EventKind.Information:
                    case EventKind.Verbose:
                        Trace.TraceInformation(message, args);
                        break;
                }
            }
        }
    }

    internal class BlobProvider
    {
        private StorageAccountInfo _info;
        private BlobContainer _container;
        private string _containerName;
        private object _lock = new object();

        private static readonly TimeSpan _Timeout = TimeSpan.FromSeconds(30);
        private static readonly RetryPolicy _RetryPolicy = RetryPolicies.RetryN(3, TimeSpan.FromSeconds(1));
        private const string _PathSeparator = "/";


        internal BlobProvider(StorageAccountInfo info, string containerName)
        {
            this._info = info;
            this._containerName = containerName;
        }

        internal string ContainerUrl
        {
            get
            {
                return string.Join(_PathSeparator, new string[] { _info.BaseUri.AbsolutePath, _containerName });
            }
        }

        internal bool GetBlobContentsWithoutInitialization(string blobName, Stream outputStream, out BlobProperties properties)
        {
            Debug.Assert(outputStream != null);

            BlobContainer container = GetContainer();

            try
            {
                properties = container.GetBlob(blobName, new BlobContents(outputStream), false);
                Log.Write(EventKind.Information, "Getting contents of blob {0}", _info.BaseUri + _PathSeparator + _containerName + _PathSeparator + blobName);
                return true;
            }
            catch (StorageClientException sc)
            {
                if (sc.ErrorCode == StorageErrorCode.ResourceNotFound || sc.ErrorCode == StorageErrorCode.BlobNotFound)
                {
                    properties = null;
                    return false;
                }
                else
                    throw;
            }
        }

        internal MemoryStream GetBlobContent(string blobName, out BlobProperties properties)
        {
            MemoryStream blobContent = new MemoryStream();
            properties = GetBlobContent(blobName, blobContent);
            blobContent.Seek(0, SeekOrigin.Begin);
            return blobContent;
        }


        internal BlobProperties GetBlobContent(string blobName, Stream outputStream)
        {
            BlobProperties properties;
            BlobContainer container = GetContainer();
            try
            {
                properties = container.GetBlob(blobName, new BlobContents(outputStream), false);
                Log.Write(EventKind.Information, "Getting contents of blob {0}", ContainerUrl + _PathSeparator + blobName);
                return properties;
            }
            catch (StorageClientException sc)
            {
                Log.Write(EventKind.Error, "Error getting contents of blob {0}: {1}", ContainerUrl + _PathSeparator + blobName, sc.Message);
                throw;
            }
        }

        internal void UploadStream(string blobName, Stream output)
        {
            UploadStream(blobName, output, true);
        }

        internal bool UploadStream(string blobName, Stream output, bool overwrite)
        {
            BlobContainer container = GetContainer();
            try
            {
                output.Position = 0; //Rewind to start
                Log.Write(EventKind.Information, "Uploading contents of blob {0}", ContainerUrl + _PathSeparator + blobName);
                BlobProperties properties = new BlobProperties(blobName);
                return container.CreateBlob(properties, new BlobContents(output), overwrite);
            }
            catch (StorageException se)
            {
                Log.Write(EventKind.Error, "Error uploading blob {0}: {1}", ContainerUrl + _PathSeparator + blobName, se.Message);
                throw;
            }
        }

        internal bool DeleteBlob(string blobName)
        {
            BlobContainer container = GetContainer();
            try
            {
                return container.DeleteBlob(blobName);
            }
            catch (StorageException se)
            {
                Log.Write(EventKind.Error, "Error deleting blob {0}: {1}", ContainerUrl + _PathSeparator + blobName, se.Message);
                throw;
            }
        }

        internal bool DeleteBlobsWithPrefix(string prefix)
        {
            bool ret = true;

            IEnumerable<BlobProperties> e = ListBlobs(prefix);
            if (e == null)
            {
                return true;
            }
            IEnumerator<BlobProperties> props = e.GetEnumerator();
            if (props == null)
            {
                return true;
            }
            while (props.MoveNext())
            {
                if (props.Current != null)
                {
                    if (!DeleteBlob(props.Current.Name))
                    {
                        // ignore this; it is possible that another thread could try to delete the blob
                        // at the same time
                        ret = false;
                    }
                }
            }
            return ret;
        }

        public IEnumerable<BlobProperties> ListBlobs(string folder)
        {
            BlobContainer container = GetContainer();
            try
            {
                return container.ListBlobs(folder, false).OfType<BlobProperties>();
            }
            catch (StorageException se)
            {
                Log.Write(EventKind.Error, "Error enumerating contents of folder {0} exists: {1}", ContainerUrl + _PathSeparator + folder, se.Message);
                throw;
            }
        }

        private BlobContainer GetContainer()
        {
            // we have to make sure that only one thread tries to create the container
            lock (_lock)
            {
                if (_container != null)
                {
                    return _container;
                }
                try
                {
                    BlobContainer container = BlobStorage.Create(_info).GetBlobContainer(_containerName);
                    container.Timeout = _Timeout;
                    container.RetryPolicy = _RetryPolicy;
                    container.CreateContainer();
                    _container = container;
                    return _container;
                }
                catch (StorageException se)
                {
                    Log.Write(EventKind.Error, "Error creating container {0}: {1}", ContainerUrl, se.Message);
                    throw;
                }
            }
        }

    }
}