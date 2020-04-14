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
// <copyright file="Queue.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
//
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Configuration;
using System.Collections.Specialized;
using System.Xml;
using System.IO;
using System.Globalization;
using System.Diagnostics.CodeAnalysis;


// disable the generation of warnings for missing documentation elements for 
// public classes/members in this file
#pragma warning disable 1591


namespace Microsoft.Samples.ServiceHosting.StorageClient
{

    /// <summary>
    /// The entry point of the queue storage API
    /// </summary>
    public abstract class QueueStorage
    {
        /// <summary>
        /// Factory method for QueueStorage
        /// </summary>
        /// <param name="baseUri">The base URI of the queue service</param>
        /// <param name="usePathStyleUris">If true, path-style URIs (http://baseuri/accountname/containername/objectname) are used.
        /// If false host-style URIs (http://accountname.baseuri/containername/objectname) are used,
        /// where baseuri is the URI of the service.
        /// If null, the choice is made automatically: path-style URIs if host name part of base URI is an 
        /// IP addres, host-style otherwise.</param>
        /// <param name="accountName">The name of the storage account</param>
        /// <param name="base64Key">Authentication key used for signing requests</param>
        /// <returns>A newly created QueueStorage instance</returns>
        public static QueueStorage Create(Uri baseUri, bool? usePathStyleUris, string accountName, string base64Key)
        {
            return new QueueStorageRest(new StorageAccountInfo(baseUri, usePathStyleUris, accountName, base64Key),null);
        }

        public static QueueStorage Create(StorageAccountInfo accountInfo)
        {
            return new QueueStorageRest(accountInfo, null);
        }


        public static QueueStorage Create(StorageAccountInfo accountInfo,string version)
        {
            return new QueueStorageRest(accountInfo,version);
        }

        /// <summary>
        /// Get a reference to a Queue object with a specified name. This method does not make a call to
        /// the queue service.
        /// </summary>
        /// <param name="queueName">The name of the queue</param>
        /// <returns>A newly created queue object</returns>
        public abstract MessageQueue GetQueue(string queueName);


        /// <summary>
        /// Lists the queues within the account.
        /// </summary>
        /// <returns>A list of queues</returns>
        public virtual IEnumerable<MessageQueue> ListQueues()
        {
            return ListQueues(null);
        }

        /// <summary>
        /// Lists the queues within the account that start with the given prefix.
        /// </summary>
        /// <param name="prefix">If prefix is null returns all queues.</param>
        /// <returns>A list of queues.</returns>
        public abstract IEnumerable<MessageQueue> ListQueues(string prefix);

        /// <summary>
        /// The time out for each request to the storage service.
        /// </summary>
        public TimeSpan Timeout
        {
            get;
            set;
        }

        /// <summary>
        /// The retry policy used for retrying requests
        /// </summary>
        public RetryPolicy RetryPolicy
        {
            get;
            set;
        }

        /// <summary>
        /// The base URI of the blob storage service
        /// </summary>
        public Uri BaseUri
        {
            get
            {
                return AccountInfo.BaseUri;
            }
        }

        /// <summary>
        /// The name of the storage account
        /// </summary>
        public string AccountName
        {
            get
            {
                return AccountInfo.AccountName;
            }
        }

        /// <summary>
        /// Indicates whether to use/generate path-style or host-style URIs
        /// </summary>
        public bool UsePathStyleUris
        {
            get
            {
                return AccountInfo.UsePathStyleUris;
            }
        }

        /// <summary>
        /// The default timeout
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes",
                          Justification = "TimeSpan is a non-mutable type")]
        public static readonly TimeSpan DefaultTimeout = TimeSpan.FromSeconds(30);


        /// <summary>
        /// The default retry policy
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes",
                          Justification = "RetryPolicy is a non-mutable type")]
        public static readonly RetryPolicy DefaultRetryPolicy = RetryPolicies.NoRetry;

        internal protected QueueStorage(StorageAccountInfo accountInfo,string version)
        {
            this.AccountInfo = accountInfo;
            Timeout = DefaultTimeout;
            RetryPolicy = DefaultRetryPolicy;
            Version = version;
        }

        protected QueueStorage(QueueStorage other)
        {
            this.AccountInfo = other.AccountInfo;
            this.Timeout = other.Timeout;
            this.RetryPolicy = other.RetryPolicy;
            this.Version = other.Version;
        }

        internal protected StorageAccountInfo AccountInfo { get; set; }
        internal protected SharedKeyCredentials Credentials { get; set; }
        internal protected string Version { get; set; }
    }

    /// <summary>
    /// Objects of this class represent a single message in the queue.
    /// </summary>
    public class Message
    {

        /// <summary>
        /// The maximum message size in bytes.
        /// </summary>
        public static readonly int MaxMessageSize = 8 * 1024;

        /// <summary>
        /// The maximum amount of time a message is kept in the queue. Max value is 7 days. 
        /// Value is given in seconds.
        /// </summary>
        public static readonly int MaxTimeToLive = 7 * 24 * 60 * 60;

        /// <summary>
        /// This constructor is not publicly exposed.
        /// </summary>
        internal Message()
        {
        }

        /// <summary>
        /// Creates a message and initializes the content of the message to be the specified string.
        /// </summary>
        /// <param name="content">A string representing the contents of the message.</param>
        public Message(string content)
        {
            if (content == null)
            {
                throw new ArgumentNullException("content");
            }
            this.content = Encoding.UTF8.GetBytes(content);
        }

        /// <summary>
        /// Creates a message and given the specified byte contents.
        /// In this implementation, regardless of whether an XML or binary data is passed into this 
        /// function, message contents are converted to base64 before passing the data to the queue service.
        /// When calculating the size of the message, the size of the base64 encoding is thus the important
        /// parameter.
        /// </summary>
        /// <param name="content"></param>
        public Message(byte[] content)
        {
            if (content == null)
            {
                throw new ArgumentNullException("content");
            }
            if (Convert.ToBase64String(content).Length > MaxMessageSize)
            {
                throw new ArgumentException("Message body is too big!");
            }
            this.content = content;
        }


        /// <summary>
        /// A unique ID of the message as returned from queue operations.
        /// </summary>
        public string Id
        {
            get;
            internal set;
        }

        /// <summary>
        /// When a message is retrieved from a queue, a PopReceipt is returned. The PopReceipt is used when 
        /// deleting a message from the queue.
        /// </summary>
        public string PopReceipt
        {
            get;
            internal set;
        }

        /// <summary>
        /// The point in time when the message was put into the queue.
        /// </summary>
        public DateTime InsertionTime
        {
            get;
            internal set;
        }

        /// <summary>
        /// A message's expiration time.
        /// </summary>
        public DateTime ExpirationTime
        {
            get;
            internal set;
        }

        /// <summary>
        /// The point in time when a message becomes visible again after a Get() operation was called 
        /// that returned the message.
        /// </summary>
        public DateTime TimeNextVisible
        {
            get;
            internal set;
        }

        /// <summary>
        /// Returns the the contents of the message as a string.
        /// </summary>
        public string ContentAsString()
        {
            return Encoding.UTF8.GetString(this.content);
        }

        /// <summary>
        /// Returns the content of the message as a byte array
        /// </summary>
        public byte[] ContentAsBytes()
        {
            return content;
        }

        /// <summary>
        /// When calling the Get() operation on a queue, the content of messages 
        /// returned in the REST protocol are represented as Base64-encoded strings.
        /// This internal function transforms the Base64 representation into a byte array.
        /// </summary>
        /// <param name="str">The Base64-encoded string.</param>
        internal void SetContentFromBase64String(string str) {
            if (str == null || str == string.Empty)
            {
                // we got a message with an empty <MessageText> element
                this.content = Encoding.UTF8.GetBytes(string.Empty);
            }
            else
            {
                this.content = Convert.FromBase64String(str);
            }
        }

        /// <summary>
        /// Internal method used for creating the XML that becomes part of a REST request
        /// </summary>
        internal byte[] GetContentXMLRepresentation(out int length)
        {
            length = 0;
            byte[] ret = null;
            StringBuilder builder = new StringBuilder();            
            XmlWriterSettings settings = new XmlWriterSettings();
            settings.OmitXmlDeclaration = true;            
            System.Text.UTF8Encoding enc = new UTF8Encoding(false);            
            using (XmlWriter writer = XmlWriter.Create(builder, settings))
            {
                writer.WriteStartDocument();
                writer.WriteStartElement(StorageHttpConstants.XmlElementNames.QueueMessage);
                writer.WriteStartElement(StorageHttpConstants.XmlElementNames.MessageText);
                writer.WriteRaw(Convert.ToBase64String(content));
                writer.WriteEndElement();
                writer.WriteEndElement();
                writer.WriteEndDocument();
                writer.Flush();
            }
            ret = enc.GetBytes(builder.ToString());          
            if (ret != null)
            {
                length = ret.Length;
            }
            return ret;
        }        

        private byte[] content;
    }

    /// <summary>
    /// Queues in the storage client library expose a functionality for listening for incoming messages. 
    /// If a message is put into a queue, a corresponding event is issued and this delegate is called. This functionality
    /// is implemented internally in this library by periodically polling for incoming messages.
    /// </summary>
    /// <param name="sender">The queue that has received a new event.</param>
    /// <param name="e">The event argument containing the message.</param>
    public delegate void MessageReceivedEventHandler(object sender, MessageReceivedEventArgs e);

    /// <summary>
    /// The argument class for the MessageReceived event.
    /// </summary>
    public class MessageReceivedEventArgs : EventArgs
    {
        /// <summary>
        /// The message itself.
        /// </summary>
        private Message _msg;

        /// <summary>
        /// Constructor for creating a message received argument.
        /// </summary>
        /// <param name="msg"></param>
        public MessageReceivedEventArgs(Message msg)
        {
            if (msg == null)
            {
                throw new ArgumentNullException("msg");
            }
            _msg = msg;
        }

        /// <summary>
        /// The message received by the queue.
        /// </summary>
        public Message Message
        {
            get
            {
                return _msg;
            }
            set
            {
                if (value == null)
                {
                    throw new ArgumentNullException("value");
                }
                _msg = value;
            }
        }
    }

    public class QueueProperties
    {
        /// <summary>
        /// The approximated amount of messages in the queue.
        /// </summary>
        public int ApproximateMessageCount
        {
            get;
            internal set;
        }

        /// <summary>
        /// Metadata for the queue in the form of name-value pairs.
        /// </summary>
        public NameValueCollection Metadata
        {
            get;
            set;
        }
    }

    /// <summary>
    /// Objects of this class represent a queue in a user's storage account.
    /// </summary>
    public abstract class MessageQueue
    {

        /// <summary>
        /// The name of the queue.
        /// </summary>
        private string _name;

        /// <summary>
        /// The user account this queue lives in.
        /// </summary>
        private StorageAccountInfo _account;


        /// <summary>
        /// This constructor is only called by subclasses.
        /// </summary>
        internal protected MessageQueue()
        {
            // queues are generated using factory methods
        }

        internal protected MessageQueue(string name, StorageAccountInfo account)
        {
            if (string.IsNullOrEmpty(name))
            {
                throw new ArgumentException("Queue name cannot be null or empty!");
            }
            if (account == null)
            {
                throw new ArgumentNullException("account");
            }
            if (!Utilities.IsValidContainerOrQueueName(name))
            {
                throw new ArgumentException(string.Format(CultureInfo.InvariantCulture, "The specified queue name \"{0}\" is not valid!" +
                            "Please choose a name that conforms to the naming conventions for queues!", name));
            }
            _name = name;
            _account = account;
        }

        /// <summary>
        /// The name of the queue exposed as a public property.
        /// </summary>
        public string Name
        {
            get
            {
                return _name;
            }
        }

        /// <summary>
        /// The account info object this queue lives in -- exposed as an internal property.
        /// </summary>
        internal StorageAccountInfo AccountInfo
        {
            get {
                return _account;
            }
        }

        /// <summary>
        /// Indicates whether to use/generate path-style or host-style URIs
        /// </summary>
        public bool UsePathStyleUris
        {
            get
            {
                return _account.UsePathStyleUris;
            }
        }

        /// <summary>
        /// The URI of the queue
        /// </summary>
        public abstract Uri QueueUri
        {
            get;
        }

        /// <summary>
        /// The retry policy used for retrying requests; this is the retry policy of the 
        /// storage account where this queue was created
        /// </summary>
        public RetryPolicy RetryPolicy
        {
            get;
            set;
        }

        /// <summary>
        /// The timeout of requests.
        /// </summary>
        public TimeSpan Timeout
        {
            get;
            set;
        }

        /// <summary>
        /// Creates a queue in the specified storage account.
        /// </summary>
        /// <param name="queueAlreadyExists">true if a queue with the same name already exists.</param>
        /// <returns>true if the queue was successfully created.</returns>
        public abstract bool CreateQueue(out bool queueAlreadyExists);

        /// <summary>
        /// Creates a queue in the specified storage account.
        /// </summary>
        /// <returns>true if the queue was successfully created.</returns>
        public virtual bool CreateQueue()
        {
            bool ignore;
            return CreateQueue(out ignore);
        }


        /// <summary>
        /// Determines whether a queue with the same name already exists in an account.
        /// </summary>
        /// <returns>true if a queue with the same name already exists.</returns>
        public virtual bool DoesQueueExist()
        {
            try
            {
                this.GetProperties();
                return true;
            }
            catch (StorageClientException e)
            {
                if (e.ErrorCode == StorageErrorCode.ResourceNotFound)
                    return false;
                throw;
            }
        }

        /// <summary>
        /// Deletes the queue. The queue will be deleted regardless of whether there are messages in the 
        /// queue or not.
        /// </summary>
        /// <returns>true if the queue was successfully deleted.</returns>
        public abstract bool DeleteQueue();

        /// <summary>
        /// Sets the properties of a queue.
        /// </summary>
        /// <param name="properties">The queue's properties to set.</param>
        /// <returns>true if the properties were successfully written to the queue.</returns>
        public abstract bool SetProperties(QueueProperties properties);

        /// <summary>
        /// Retrieves the queue's properties.
        /// </summary>
        /// <returns>The queue's properties.</returns>
        public abstract QueueProperties GetProperties();

        /// <summary>
        /// Retrieves the approximate number of messages in a queue.
        /// </summary>
        /// <returns>The approximate number of messages in this queue.</returns>
        public abstract int ApproximateCount();

        /// <summary>
        /// Puts a message in the queue.
        /// </summary>
        /// <param name="msg">The message to store in the queue.</param>
        /// <returns>true if the message has been successfully enqueued.</returns>
        public abstract bool PutMessage(Message msg);

        /// <summary>
        /// Puts a message in the queue.
        /// </summary>
        /// <param name="msg">The message to store in the queue.</param>
        /// <param name="timeToLiveInSeconds">The time to live for the message in seconds.</param>
        /// <returns>true if the message has been successfully enqueued.</returns>
        public abstract bool PutMessage(Message msg, int timeToLiveInSeconds);

        /// <summary>
        /// Retrieves a message from the queue. 
        /// </summary>
        /// <returns>The message retrieved or null if the queue is empty.</returns>
        public abstract Message GetMessage();

        /// <summary>
        /// Retrieves a message and sets its visibility timeout to the specified number of seconds.
        /// </summary>
        /// <param name="visibilityTimeoutInSeconds">Visibility timeout of the message retrieved in seconds.</param>
        /// <returns></returns>
        public abstract Message GetMessage(int visibilityTimeoutInSeconds);

        /// <summary>
        /// Tries to retrieve the given number of messages.
        /// </summary>
        /// <param name="numberOfMessages">Maximum number of messages to retrieve.</param>
        /// <returns>The list of messages retrieved.</returns>
        public abstract IEnumerable<Message> GetMessages(int numberOfMessages);

        /// <summary>
        /// Tries to retrieve the given number of messages.
        /// </summary>
        /// <param name="numberOfMessages">Maximum number of messages to retrieve.</param>
        /// <param name="visibilityTimeoutInSeconds">The visibility timeout of the retrieved messages in seconds.</param>
        /// <returns>The list of messages retrieved.</returns>
        public abstract IEnumerable<Message> GetMessages(int numberOfMessages, int visibilityTimeoutInSeconds);

        /// <summary>
        /// Get a message from the queue but do not actually dequeue it. The message will remain visible 
        /// for other parties requesting messages.
        /// </summary>
        /// <returns>The message retrieved or null if there are no messages in the queue.</returns>
        public abstract Message PeekMessage();

        /// <summary>
        /// Tries to get a copy of messages in the queue without actually dequeuing the messages.
        /// The messages will remain visible in the queue.
        /// </summary>
        /// <param name="numberOfMessages">Maximum number of message to retrieve.</param>
        /// <returns>The list of messages retrieved.</returns>
        public abstract IEnumerable<Message> PeekMessages(int numberOfMessages);

        /// <summary>
        /// Deletes a message from the queue.
        /// </summary>
        /// <param name="msg">The message to retrieve with a valid popreceipt.</param>
        /// <returns>true if the operation was successful.</returns>
        public abstract bool DeleteMessage(Message msg);

        /// <summary>
        /// Delete all messages in a queue.
        /// </summary>
        /// <returns>true if all messages were deleted successfully.</returns>
        public abstract bool Clear();

        /// <summary>
        /// The default time interval between polling the queue for messages. 
        /// Polling is only enabled if the user has called StartReceiving().
        /// </summary>
        public static readonly int DefaultPollInterval = 5000;

        /// <summary>
        /// The poll interval in milliseconds. If not explicitly set, this defaults to 
        /// the DefaultPollInterval.
        /// </summary>
        public abstract int PollInterval
        {
            get;
            set;
        }

        /// <summary>
        /// Starts the automatic reception of messages.
        /// </summary>
        /// <returns>true if the operation was successful.</returns>
        public abstract bool StartReceiving();

        /// <summary>
        /// Stop the automatic reception of messages.
        /// </summary>
        public abstract void StopReceiving();

        /// <summary>
        /// The event users subscribe to in order to automatically receive messages
        /// from a queue.
        /// </summary>
        public abstract event MessageReceivedEventHandler MessageReceived;

    }

}