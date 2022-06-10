using System;
using System.Configuration; // Namespace for ConfigurationManager
using Azure.Storage.Queues; // Namespace for Queue storage types

namespace TeamsAuth
{
    public class SimpleStorageAccountClient
    {
        string storageConnectionString = "DefaultEndpointsProtocol=https;AccountName=jostoragesample;AccountKey=id8WyYPkC3ekXGOKDU1CAOk/HX55by8PjiZQ+RsMqdxLjrlyPReGG/ZsCGRjfp0pU0ZePMCQ21o9+AStgU2bdQ==;EndpointSuffix=core.windows.net";

        public void InsertMessage(string queueName, string message)
        {
            // Get the connection string from app settings
            string connectionString = ConfigurationManager.AppSettings[storageConnectionString];

            // Instantiate a QueueClient which will be used to create and manipulate the queue
            QueueClient queueClient = new QueueClient(storageConnectionString, queueName);

            // Create the queue if it doesn't already exist
            queueClient.CreateIfNotExists();

            if (queueClient.Exists())
            {
                // format the message so the queue accepts it
                var plainTextBytesMessage = System.Text.Encoding.UTF8.GetBytes(message);

                // Send a message to the queue
                queueClient.SendMessage(System.Convert.ToBase64String(plainTextBytesMessage));
            }

            Console.WriteLine($"Inserted: {message}");
        }
    }
}
