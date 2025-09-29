from rest_framework.routers import DefaultRouter
from .views import EventsHandlerViewSet

router = DefaultRouter()
router.register(r'events', EventsHandlerViewSet)

urlpatterns = router.urls
