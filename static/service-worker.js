self.addEventListener("install", function(event) {
    self.skipWaiting();
});

self.addEventListener("activate", function(event) {
    event.waitUntil(self.clients.claim());
});

self.addEventListener("push", function(event) {
    let data = {
        title: "塾代講管理",
        body: "通知があります"
    };

    if (event.data) {
        data = event.data.json();
    }

    event.waitUntil(
        self.registration.showNotification(data.title, {
            body: data.body,
            icon: "/static/icon-192.png",
            badge: "/static/icon-192.png",
            vibrate: [200, 100, 200]
        })
    );
});

self.addEventListener("notificationclick", function(event) {
    event.notification.close();

    event.waitUntil(
        clients.openWindow("/")
    );
});