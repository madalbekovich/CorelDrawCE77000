from django.contrib import admin
from .models import EventsHandler
from django.utils.safestring import mark_safe

@admin.register(EventsHandler)
class EventsHandlerAdmin(admin.ModelAdmin):
    list_display = ("id", "title", "created_at", "plotter", "get_preview" )
    list_filter = ("status", 'created_at')

    def get_preview(self, obj):
        if obj.preview:
            return mark_safe(f'<img src="{obj.preview.url}" width="150" />')
        return "Нет превью"


