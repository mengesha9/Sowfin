using Sowfin.API.Notifications.Abstraction;

namespace Sowfin.API.Notifications.Models
{
  public class Notification<T> : INotification
  {
    public NotificationType NotificationType { get; set; }
    public T Payload { get; set; }

  }
}