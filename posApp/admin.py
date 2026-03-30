from django.contrib import admin
from posApp.models import Category, Products, Sales, salesItems, Customer

# Register your models here.
admin.site.register(Category)
admin.site.register(Products)
admin.site.register(Sales)
admin.site.register(salesItems)
admin.site.register(Customer)
# admin.site.register(Report)
# admin.site.register(Employees)
