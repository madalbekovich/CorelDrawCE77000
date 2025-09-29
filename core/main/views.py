from rest_framework import viewsets
from .models import EventsHandler
from .serializers import EventsHandlerSerializer

class EventsHandlerViewSet(viewsets.ModelViewSet):
    queryset = EventsHandler.objects.all()
    serializer_class = EventsHandlerSerializer

    def create(self, request, *args, **kwargs):
        print("FILES:", request.FILES)
        print("DATA:", request.data)
        return super().create(request, *args, **kwargs)