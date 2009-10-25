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
// <copyright file="RestQueue.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
//
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;
using System.Net;
using System.Collections.Specialized;
using System.IO;
using System.Xml;
using System.Diagnostics;
using System.Globalization;
using System.Threading;


namespace Microsoft.Samples.ServiceHosting.StorageClient
{

    internal class QueueStorageRest : QueueStorage
    {
        private SharedKeyCredentials _credentials;

        internal QueueStorageRest(StorageAccountInfo accountInfo,string version)
            : base(accountInfo,version)
        {
            byte[] key = null;
            if (accountInfo.Base64Key != null)
            {
                key = Convert.FromBase64String(accountInfo.Base64Key);
            }
            _credentials = new SharedKeyCredentials(accountInfo.AccountName, key);
        }

        /// <summary>
        /// Get a reference to a Queue object with a specified name. This method does not make a call to
        /// the queue service.
        /// </summary>
        /// <param name="queueName">The name of the queue</param>
        /// <returns>A newly created queue object</returns>
        public override MessageQueue GetQueue(string queueName)
        {
            return new QueueRest(queueName,
                                 AccountInfo, 
                                 Timeout,
                                 RetryPolicy,
                                 Version
                                 );
        }

        internal class ListQueueResult
        {
            internal ListQueueResult(IEnumerable<string> names, IEnumerable<string> urls, string nextMarker)
            {
                Names = names;
                Urls = urls;
                NextMarker = nextMarker;
            }

            internal IEnumerable<string> Names
            {
                get;
                private set;
            }

            internal IEnumerable<string> Urls
            {
                get;
                private set;
            }

            internal string NextMarker
            {
                get;
                private set;
            }
        }

        /// <summary>
        /// Lists all queues with a given prefix within an account.
        /// </summary>
        /// <param name="prefix"></param>
        /// <returns>The list of queue names.</returns>
        public override IEnumerable<MessageQueue> ListQueues(string prefix)
        {
            string marker = "";
            const int maxResults = StorageHttpConstants.ListingConstants.MaxQueueListResults;

            do
            {
                ListQueueResult result = ListQueuesImpl(prefix, marker, maxResults);
                if (result == null)
                {
                    marker = null;
                }
                else
                {
                    marker = result.NextMarker;

                    foreach (string name in result.Names)
                    {
                        yield return new QueueRest(name, AccountInfo, this.Timeout, this.RetryPolicy,this.Version);
                    }
                }
            } while (marker != null);
        }

        /// <summary>
        /// Lists the queues within the account.
        /// </summary>
        /// <returns>A list of queues</returns>
        private ListQueueResult ListQueuesImpl(string prefix, string marker, int maxResult)
        {
            ListQueueResult result = null;

            RetryPolicy(() =>
            {
                NameValueCollection col = new NameValueCollection();
                col.Add(StorageHttpConstants.QueryParams.QueryParamComp, StorageHttpConstants.CompConstants.List);
                if (!string.IsNullOrEmpty(prefix))
                {
                    col.Add(StorageHttpConstants.QueryParams.QueryParamPrefix, prefix);
                }
                if (!string.IsNullOrEmpty(marker))
                {
                    col.Add(StorageHttpConstants.QueryParams.QueryParamMarker, marker);
                }
                col.Add(StorageHttpConstants.QueryParams.QueryParamMaxResults, maxResult.ToString(CultureInfo.InvariantCulture));

                ResourceUriComponents uriComponents;
                Uri uri = Utilities.CreateRequestUri(
                                        AccountInfo.BaseUri,
                                        AccountInfo.UsePathStyleUris,
                                        AccountInfo.AccountName,
                                        null,
                                        null,
                                        Timeout,
                                        col,
                                        out uriComponents
                                        );
                HttpWebRequest request = Utilities.CreateHttpRequest(uri, StorageHttpConstants.HttpMethod.Get, Timeout);
                _credentials.SignRequest(request, uriComponents);

                try
                {
                    using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                    {
                        if (response.StatusCode == HttpStatusCode.OK)
                        {
                            using (Stream stream = response.GetResponseStream())
                            {
                                result = GetQueuesFromResponse(stream);
                                stream.Close();
                            }
                        }
                        else
                        {
                            Utilities.ProcessUnexpectedStatusCode(response);
                        }
                        response.Close();
                    }
                }
                catch (WebException we)
                {
                    throw Utilities.TranslateWebException(we);
                }
            });

            return result;
        }

        private static ListQueueResult GetQueuesFromResponse(Stream stream)
        {
            ListQueueResult result = null;
            List<string> names = new List<string>();
            List<string> urls = new List<string>();
            string nextMarker = null;

            XmlDocument doc = new XmlDocument();
            try
            {
                doc.Load(stream);
            }
            catch (XmlException xe)
            {
                throw new StorageServerException(StorageErrorCode.ServiceBadResponse,
                    "The result of a ListQueue operation could not be parsed", default(HttpStatusCode), xe);
            }

            // get queue names and urls
            XmlNodeList queueNameNodes = doc.SelectNodes(XPathQueryHelper.QueueListQuery);
            foreach (XmlNode queueNameNode in queueNameNodes)
            {
                string queueName = XPathQueryHelper.LoadSingleChildStringValue(queueNameNode, StorageHttpConstants.XmlElementNames.QueueName, true);
                names.Add(queueName);
                string url = XPathQueryHelper.LoadSingleChildStringValue(queueNameNode, StorageHttpConstants.XmlElementNames.Url, true);
                urls.Add(url);
            }

            // Get the nextMarker
            XmlNode nextMarkerNode = doc.SelectSingleNode(XPathQueryHelper.NextMarkerQuery);
            if (nextMarkerNode != null && nextMarkerNode.FirstChild != null)
            {
                nextMarker = nextMarkerNode.FirstChild.Value;
            }
            if (names.Count > 0)
            {
                Debug.Assert(names.Count == urls.Count);
                result = new ListQueueResult(names, urls, nextMarker);
            }
            return result;
        }
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1001:TypesThatOwnDisposableFieldsShouldBeDisposable", 
                                                     Justification = "Disposable types are only used for automatic receiving of messages, which handle the disposal themselves.")]
    internal class QueueRest : MessageQueue
    {
        #region Member variables and constructors

        private Uri _queueUri;
        private SharedKeyCredentials _credentials;
        private int _pollInterval = DefaultPollInterval;
        private string version;


        internal QueueRest(
                    string name,
                    StorageAccountInfo account,
                    TimeSpan timeout,
                    RetryPolicy retryPolicy,
                    string version
                    )
            : base(name, account)
        {
            byte[] key = null;
            if (AccountInfo.Base64Key != null)
            {
                key = Convert.FromBase64String(AccountInfo.Base64Key);
            }
            ResourceUriComponents uriComponents = new ResourceUriComponents(account.AccountName, name, null);
            _credentials = new SharedKeyCredentials(AccountInfo.AccountName, key);
            _queueUri = HttpRequestAccessor.ConstructResourceUri(account.BaseUri, uriComponents, account.UsePathStyleUris);
            Timeout = timeout;
            RetryPolicy = retryPolicy;
            this.version = version;
        }
        #endregion

        #region Public interface

        public override Uri QueueUri
        {
            get
            {
                return _queueUri;
            }
        }

        public override bool CreateQueue(out bool queueAlreadyExists)
        {
            bool result = false;
            queueAlreadyExists = false;
            // cannot use ref or out parameters in the retry expression below
            bool exists = false;

            RetryPolicy(() =>
            {
                ResourceUriComponents uriComponents;
                Uri uri = CreateRequestUri(null, new NameValueCollection(), false, out uriComponents);
                HttpWebRequest request = CreateHttpRequest(uri, StorageHttpConstants.HttpMethod.Put);
                _credentials.SignRequest(request, uriComponents);

                try
                {
                    using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                    {
                        if (response.StatusCode == HttpStatusCode.Created)
                        {
                            // as an addition we could parse the result and retrieve
                            // queue properties at this point
                            exists = false;
                            result = true;
                        }
                        else if (response.StatusCode == HttpStatusCode.NoContent)
                        {
                            exists = true;
                        }
                        else
                        {
                            Utilities.ProcessUnexpectedStatusCode(response);
                        }
                        response.Close();
                    }
                }
                catch (WebException we)
                {
                    if (we.Response != null && 
                        ((HttpWebResponse)we.Response).StatusCode == HttpStatusCode.Conflict)
                    {
                        exists = true;
                    }
                    else
                    {
                        throw Utilities.TranslateWebException(we);
                    }
                }
            });

            queueAlreadyExists = exists;
            return result;
        }

        public override bool DeleteQueue()
        {
            bool result = false;

            RetryPolicy(() =>
            {
                NameValueCollection col = new NameValueCollection();                
                
                ResourceUriComponents uriComponents;
                Uri uri = CreateRequestUri(null, col, false, out uriComponents);
                HttpWebRequest request = CreateHttpRequest(uri, StorageHttpConstants.HttpMethod.Delete);
                _credentials.SignRequest(request, uriComponents);

                try
                {
                    using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                    {
                        if (response.StatusCode == HttpStatusCode.NoContent)
                        {
                            result = true;
                        }
                        else
                        {
                            Utilities.ProcessUnexpectedStatusCode(response);
                        }
                        response.Close();
                    }
                }
                catch (WebException we)
                {
                    if (we.Response != null &&
                        (((HttpWebResponse)we.Response).StatusCode == HttpStatusCode.NotFound ||
                         ((HttpWebResponse)we.Response).StatusCode == HttpStatusCode.PreconditionFailed ||
                         ((HttpWebResponse)we.Response).StatusCode == HttpStatusCode.Conflict))
                    {
                        result = false;
                    }
                    else
                    {
                        throw Utilities.TranslateWebException(we);
                    }
                }
            });

            return result;
        }

        public override bool SetProperties(QueueProperties properties)
        {
            if (properties == null)
            {
                throw new ArgumentNullException("properties");
            }

            bool result = false;

            RetryPolicy(() =>
            {
                NameValueCollection col = new NameValueCollection();
                col.Add(StorageHttpConstants.QueryParams.QueryParamComp, StorageHttpConstants.CompConstants.Metadata);

                ResourceUriComponents uriComponents;
                Uri uri = CreateRequestUri(null, col, out uriComponents);
                HttpWebRequest request = CreateHttpRequest(uri, StorageHttpConstants.HttpMethod.Put, properties.Metadata);
                _credentials.SignRequest(request, uriComponents);

                try
                {
                    using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                    {
                        if (response.StatusCode == HttpStatusCode.NoContent)
                        {
                            result = true;
                        }
                        else
                        {
                            Utilities.ProcessUnexpectedStatusCode(response);
                        }
                        response.Close();
                    }
                }
                catch (WebException we)
                {
                    throw Utilities.TranslateWebException(we);
                }
            });

            return result;
        }

        public override QueueProperties GetProperties()
        {
            QueueProperties result = null;

            RetryPolicy(() =>
            {
                NameValueCollection col = new NameValueCollection();
                col.Add(StorageHttpConstants.QueryParams.QueryParamComp, StorageHttpConstants.CompConstants.Metadata);

                ResourceUriComponents uriComponents;
                Uri uri = CreateRequestUri(null, col, out uriComponents);
                HttpWebRequest request = CreateHttpRequest(uri, StorageHttpConstants.HttpMethod.Get, null);
                _credentials.SignRequest(request, uriComponents);

                try
                {
                    using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                    {
                        if (response.StatusCode == HttpStatusCode.OK)
                        {
                            result = GetPropertiesFromHeaders(response);
                        }
                        else
                        {
                            Utilities.ProcessUnexpectedStatusCode(response);
                        }
                        response.Close();
                    }
                }
                catch (WebException we)
                {
                    throw Utilities.TranslateWebException(we);
                }
            });

            return result;
        }

        public override int ApproximateCount()
        {
            QueueProperties props = GetProperties();
            return props.ApproximateMessageCount;
        }

        // getting and putting messages

        public override bool PutMessage(Message msg)
        {
            return PutMessage(msg, -1);
        }

        public override bool PutMessage(Message msg, int timeToLiveInSeconds)
        {
            if (timeToLiveInSeconds < -1)
            {
                throw new ArgumentException("ttl parameter must be equal or larger than 0.");
            }
            else if (timeToLiveInSeconds > Message.MaxTimeToLive)
            {
                throw new ArgumentException(string.Format(CultureInfo.InvariantCulture,
                    "timeToLiveHours parameter must be smaller or equal than {0}, which is 7 days in hours.", Message.MaxTimeToLive));
            }
            if (msg == null || msg.ContentAsBytes() == null)
            {
                throw new ArgumentNullException("msg");
            }
            if (Convert.ToBase64String(msg.ContentAsBytes()).Length > Message.MaxMessageSize)
            {
                throw new ArgumentException(string.Format(CultureInfo.InvariantCulture, "Messages cannot be larger than {0} bytes.", Message.MaxMessageSize));
            }

            bool result = false;

            RetryPolicy(() =>
            {
                NameValueCollection col = new NameValueCollection();
                if (timeToLiveInSeconds != -1)
                {
                    col.Add(StorageHttpConstants.RequestParams.MessageTtl, timeToLiveInSeconds.ToString(CultureInfo.InvariantCulture));
                }

                ResourceUriComponents uriComponents;
                Uri uri = CreateRequestUri(StorageHttpConstants.RequestParams.Messages, col, out uriComponents);
                HttpWebRequest request = CreateHttpRequest(uri, StorageHttpConstants.HttpMethod.Post, null);
                int len;
                byte[] body = msg.GetContentXMLRepresentation(out len);
                request.ContentLength = len;
                _credentials.SignRequest(request, uriComponents);

                try
                {
                    using (Stream requestStream = request.GetRequestStream())
                    {                                               
                        requestStream.Write(body, 0, body.Length);
                        requestStream.Close();
                        using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                        {
                            if (response.StatusCode == HttpStatusCode.Created)
                            {
                                result = true;
                            }
                            else
                            {
                                Utilities.ProcessUnexpectedStatusCode(response);
                            }
                            response.Close();
                        }
                    }
                }
                catch (WebException we)
                {
                    throw Utilities.TranslateWebException(we);
                }
            });

            return result;
        }

        public override Message GetMessage()
        {
            IEnumerable<Message> result = GetMessages(1);
            if (result == null || result.Count() == 0)
            {
                return null;
            }
            return result.First();
        }

        public override Message GetMessage(int visibilityTimeoutInSeconds)
        {
            IEnumerable<Message> result = GetMessages(1, visibilityTimeoutInSeconds);
            if (result == null || result.Count() == 0)
            {
                return null;
            }
            return result.First();
        }

        public override IEnumerable<Message> GetMessages(int numberOfMessages)
        {
            return GetMessages(numberOfMessages, -1);
        }

        public override IEnumerable<Message> GetMessages(int numberOfMessages, int visibilityTimeout)
        {
            return InternalGet(numberOfMessages, visibilityTimeout, false);
        }

        public override Message PeekMessage()
        {
            IEnumerable<Message> result = PeekMessages(1);
            if (result == null || result.Count() == 0)
            {
                return null;
            }
            return result.First();
        }

        public override IEnumerable<Message> PeekMessages(int numberOfMessages)
        {
            return InternalGet(numberOfMessages, -1, true);
        }

        // deleting messages
        public override bool DeleteMessage(Message msg)
        {
            if (msg.PopReceipt == null)
            {
                throw new ArgumentException("No PopReceipt for the given message!");
            }

            bool result = false;

            RetryPolicy(() =>
            {
                NameValueCollection col = new NameValueCollection();
                col.Add(StorageHttpConstants.RequestParams.PopReceipt, msg.PopReceipt.ToString());

                ResourceUriComponents uriComponents;
                Uri uri = CreateRequestUri(StorageHttpConstants.RequestParams.Messages + "/" + msg.Id, col, out uriComponents);
                HttpWebRequest request = CreateHttpRequest(uri, StorageHttpConstants.HttpMethod.Delete, null);
                _credentials.SignRequest(request, uriComponents);

                try
                {
                    using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                    {
                        if (response.StatusCode == HttpStatusCode.NoContent)
                        {
                            result = true;
                        }
                        else
                        {
                            Utilities.ProcessUnexpectedStatusCode(response);
                        }
                        response.Close();
                    }

                }
                catch (WebException we)
                {
                    throw Utilities.TranslateWebException(we);
                }
            });

            return result;
        }

        public override bool Clear()
        {
            bool result = false;

            RetryPolicy(() =>
            {
                NameValueCollection col = new NameValueCollection();

                ResourceUriComponents uriComponents;
                Uri uri = CreateRequestUri(StorageHttpConstants.RequestParams.Messages, col, out uriComponents);
                HttpWebRequest request = CreateHttpRequest(uri, StorageHttpConstants.HttpMethod.Delete, null);
                _credentials.SignRequest(request, uriComponents);

                try
                {
                    using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                    {
                        if (response.StatusCode == HttpStatusCode.NoContent)
                        {
                            result = true;
                        }
                        else
                        {
                            Utilities.ProcessUnexpectedStatusCode(response);
                        }
                        response.Close();
                    }

                }
                catch (WebException we)
                {
                    throw Utilities.TranslateWebException(we);
                }
            });

            return result;
        }

        // automatic receiving of messages
        public override int PollInterval
        {
            get
            {
                return _pollInterval;
            }
            set
            {
                if (value < 0)
                {
                    throw new ArgumentException("The send threshold must be a positive value.");
                }
                if (_run)
                {
                    throw new ArgumentException("You cannot set the poll interval while the receive thread is running");
                }
                _pollInterval = value;
            }
        }

        private AutoResetEvent _evStarted;
        private AutoResetEvent _evStopped;
        private AutoResetEvent _evQuit;
        private bool _run;
        private Thread _receiveThread;
        private int _internalPollInterval;


        private void PeriodicReceive()
        {
            Message msg;

            _evStarted.Set();
            _internalPollInterval = PollInterval;
            while (!_evQuit.WaitOne(_internalPollInterval, false))
            {
                // time is up, so we get the message and continue
                msg = GetMessage();
                if (msg != null)
                {
                    MessageReceived(this, new MessageReceivedEventArgs(msg));
                    // continue receiving fast until we get no message
                    _internalPollInterval = 10;
                }
                else
                {
                    // we got no message, so we can fall back to the normal speed
                    _internalPollInterval = PollInterval;
                }
            }
            _evStopped.Set();
        }

        public override bool StartReceiving()
        {
            lock (this)
            {
                if (_run)
                {
                    return true;
                }
                _run = true;
            }
            if (_evStarted == null) {
                _evStarted = new AutoResetEvent(false);
            }
            if (_evStopped == null) {
                _evStopped = new AutoResetEvent(false);
            }
            if (_evQuit == null) {
                _evQuit = new AutoResetEvent(false);
            }
            _receiveThread = new Thread(new ThreadStart(this.PeriodicReceive));
            _receiveThread.Start();
            if (!_evStarted.WaitOne(10000, false))
            {
                _receiveThread.Abort();
                CloseEvents();
                _run = false;
                return false;
            }
            return true;
        }

        public override void StopReceiving()
        {
            _evQuit.Set();
            if (!_evStopped.WaitOne(10000, false))            
            {
                _receiveThread.Abort();
            }
            CloseEvents();
            _run = false;
        }

        private void CloseEvents()
        {
            if (_evStarted != null)
            {
                _evStarted.Close();
            }
            if (_evStopped != null)
            {
                _evStopped.Close();
            }
            if (_evQuit != null)
            {
                _evQuit.Close();
            }
        }

        public override event MessageReceivedEventHandler MessageReceived;

        #endregion

        #region Helper methods

        private static QueueProperties GetPropertiesFromHeaders(HttpWebResponse response)
        {
            QueueProperties properties = new QueueProperties();
            int prefixLength = StorageHttpConstants.HeaderNames.PrefixForMetadata.Length;
            foreach (string key in response.Headers.AllKeys)
            {
                if (key.Equals(StorageHttpConstants.HeaderNames.ApproximateMessagesCount, StringComparison.OrdinalIgnoreCase))
                {
                    properties.ApproximateMessageCount = Convert.ToInt32(response.Headers[key], CultureInfo.InvariantCulture);
                }
                else if (key.StartsWith(StorageHttpConstants.HeaderNames.PrefixForMetadata, StringComparison.OrdinalIgnoreCase))
                {
                    if (properties.Metadata == null)
                    {
                        properties.Metadata = new NameValueCollection();
                    }
                    properties.Metadata.Add(key.Substring(prefixLength), response.Headers[key]);
                }
            }
            return properties;
        }


        private IEnumerable<Message> InternalGet(int numberOfMessages, int visibilityTimeout, bool peekOnly)
        {
            if (peekOnly && visibilityTimeout != -1)
            {
                throw new ArgumentException("A peek operation does not change the visibility of messages", "visibilityTimeout");
            }
            if (numberOfMessages < 1)
            {
                throw new ArgumentException("numberOfMessages must be a positive integer", "numberOfMessages");
            }
            if (visibilityTimeout < -1)
            {
                throw new ArgumentException("Visibility Timeout must be 0 or a positive integer", "visibilityTimeout");
            }

            IEnumerable<Message> result = null;

            RetryPolicy(() =>
            {
                NameValueCollection col = new NameValueCollection();
                col.Add(StorageHttpConstants.RequestParams.NumOfMessages, numberOfMessages.ToString(CultureInfo.InvariantCulture));
                if (visibilityTimeout != -1)
                {
                    col.Add(StorageHttpConstants.RequestParams.VisibilityTimeout,
                            visibilityTimeout.ToString(CultureInfo.InvariantCulture));
                }
                if (peekOnly)
                {
                    col.Add(StorageHttpConstants.RequestParams.PeekOnly,
                            peekOnly.ToString(CultureInfo.InvariantCulture));
                }

                ResourceUriComponents uriComponents;
                Uri uri = CreateRequestUri(StorageHttpConstants.RequestParams.Messages, col, out uriComponents);
                HttpWebRequest request = CreateHttpRequest(uri, StorageHttpConstants.HttpMethod.Get, null);
                _credentials.SignRequest(request, uriComponents);

                try
                {
                    using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                    {
                        if (response.StatusCode == HttpStatusCode.OK)
                        {
                            using (Stream stream = response.GetResponseStream())
                            {
                                result = GetMessageFromResponse(stream);
                                stream.Close();
                            }
                        }
                        else
                        {
                            Utilities.ProcessUnexpectedStatusCode(response);
                        }
                        response.Close();
                    }

                }
                catch (WebException we)
                {
                    throw Utilities.TranslateWebException(we);
                }
            });


            return result;
        }

        private static IEnumerable<Message> GetMessageFromResponse(Stream stream)
        {
            List<Message> result = null;
            Message msg;

            XmlDocument doc = new XmlDocument();
            try
            {
                doc.Load(stream);
            }
            catch (XmlException xe)
            {
                throw new StorageServerException(StorageErrorCode.ServiceBadResponse,
                    "The result of a get message opertation could not be parsed", default(HttpStatusCode), xe);
            }

            XmlNodeList messagesNodes = doc.SelectNodes(XPathQueryHelper.MessagesListQuery);
            if (messagesNodes.Count > 0)
            {
                result = new List<Message>();
            }
            foreach (XmlNode messageNode in messagesNodes)
            {
                msg = new Message();
                msg.Id = messageNode.SelectSingleNode(StorageHttpConstants.XmlElementNames.MessageId).FirstChild.Value.Trim();
                Debug.Assert(msg.Id != null);
                if (messageNode.SelectSingleNode(StorageHttpConstants.XmlElementNames.PopReceipt) != null)
                {
                    msg.PopReceipt = messageNode.SelectSingleNode(StorageHttpConstants.XmlElementNames.PopReceipt).FirstChild.Value.Trim();
                    Debug.Assert(msg.PopReceipt != null);
                }                
                msg.InsertionTime = XPathQueryHelper.LoadSingleChildDateTimeValue(messageNode, StorageHttpConstants.XmlElementNames.InsertionTime, false).Value;
                msg.ExpirationTime = XPathQueryHelper.LoadSingleChildDateTimeValue(messageNode, StorageHttpConstants.XmlElementNames.ExpirationTime, false).Value;
                if (XPathQueryHelper.LoadSingleChildDateTimeValue(messageNode, StorageHttpConstants.XmlElementNames.TimeNextVisible, false) != null)
                {
                    msg.TimeNextVisible = XPathQueryHelper.LoadSingleChildDateTimeValue(messageNode, StorageHttpConstants.XmlElementNames.TimeNextVisible, false).Value;
                }
                msg.SetContentFromBase64String(XPathQueryHelper.LoadSingleChildStringValue(messageNode, StorageHttpConstants.XmlElementNames.MessageText, false));
                result.Add(msg);
            }
            return result.AsEnumerable();
        }

        private HttpWebRequest CreateHttpRequest(Uri uri, string httpMethod)
        {
            return CreateHttpRequest(uri, httpMethod, null);
        }

        private HttpWebRequest CreateHttpRequest(Uri uri, string httpMethod, NameValueCollection metadata)
        {
            HttpWebRequest request = (HttpWebRequest)HttpWebRequest.Create(uri);
            request.Timeout = (int)Timeout.TotalMilliseconds;
            request.ReadWriteTimeout = (int)Timeout.TotalMilliseconds;
            request.Method = httpMethod;
            request.ContentLength = 0;
            request.Headers.Add(StorageHttpConstants.HeaderNames.StorageDateTime,
                                Utilities.ConvertDateTimeToHttpString(DateTime.UtcNow));
            if (!String.IsNullOrEmpty(this.version))
            {
                request.Headers.Add(StorageHttpConstants.HeaderNames.Version, this.version);
            }
            
            if (metadata != null)
            {
                Utilities.AddMetadataHeaders(request, metadata);
            }
            return request;
        }

        private Uri CreateRequestUri(
                        string uriSuffix,
                        NameValueCollection queryParameters,
                        out ResourceUriComponents uriComponents
                        )
        {
            return CreateRequestUri(uriSuffix, queryParameters, false, out uriComponents);
        }

        private Uri CreateRequestUri(
            string uriSuffix,
            NameValueCollection queryParameters,
            bool accountOperation,
            out ResourceUriComponents uriComponents
            )
        {
            return Utilities.CreateRequestUri(
                            this.AccountInfo.BaseUri,
                            this.AccountInfo.UsePathStyleUris,
                            this.AccountInfo.AccountName,
                            accountOperation ? null : this.Name,
                            uriSuffix,
                            this.Timeout,
                            queryParameters,
                            out uriComponents
                            );
        }

        #endregion

    }
}