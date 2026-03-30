from datetime import datetime, timedelta
from unicodedata import category
from django.db import models
from django.utils import timezone
from django.db.models import Max, F
from django.db.models.signals import pre_save, post_delete, post_save
from django.dispatch import receiver

# Create your models here.

class Employee(models.Model):
    name = models.CharField(max_length=100)
    phone_number = models.CharField(max_length=15)
    position = models.CharField(max_length=100)
    daily_wage = models.DecimalField(max_digits=10, decimal_places=2)

    def __str__(self):
        return self.name
        
class Attendance(models.Model):
    employee = models.ForeignKey(Employee, on_delete=models.CASCADE, related_name='attendances')
    date = models.DateField()
    present = models.BooleanField(default=True)
    date_added = models.DateTimeField(default=timezone.now)
    
    class Meta:
        unique_together = ['employee', 'date']
        
    def __str__(self):
        status = "Present" if self.present else "Absent"
        return f"{self.employee.name} - {self.date} ({status})"
        
    @property
    def is_working_day(self):
        # Check if the date is Tuesday (1), Wednesday (2), or Thursday (3)
        return self.date.weekday() in [1, 2, 3]
    
    @property
    def wage_for_day(self):
        if self.present and self.is_working_day:
            return self.employee.daily_wage
        return 0

class WeeklyDebit(models.Model):
    week_number = models.IntegerField()
    week_start_date = models.DateField()
    description = models.CharField(max_length=255)
    amount = models.DecimalField(max_digits=10, decimal_places=2)
    date_added = models.DateTimeField(default=timezone.now)

    def __str__(self):
        return f"Week {self.week_number}: {self.description} - ₹{self.amount}"

# class Employees(models.Model):
#     code = models.CharField(max_length=100,blank=True) 
#     firstname = models.TextField() 
#     middlename = models.TextField(blank=True,null= True) 
#     lastname = models.TextField() 
#     gender = models.TextField(blank=True,null= True) 
#     dob = models.DateField(blank=True,null= True) 
#     contact = models.TextField() 
#     address = models.TextField() 
#     email = models.TextField() 
#     department_id = models.ForeignKey(Department, on_delete=models.CASCADE) 
#     position_id = models.ForeignKey(Position, on_delete=models.CASCADE) 
#     date_hired = models.DateField() 
#     salary = models.FloatField(default=0) 
#     status = models.IntegerField() 
#     date_added = models.DateTimeField(default=timezone.now) 
#     date_updated = models.DateTimeField(auto_now=True) 

    # def __str__(self):
    #     return self.firstname + ' ' +self.middlename + ' '+self.lastname + ' '
class Category(models.Model):
    name = models.TextField()
    # description = models.TextField()
    status = models.IntegerField(default=1) 
    date_added = models.DateTimeField(default=timezone.now) 
    date_updated = models.DateTimeField(auto_now=True) 

    def __str__(self):
        return self.name

class Products(models.Model):
    category_id = models.ForeignKey(Category, on_delete=models.CASCADE)
    name = models.TextField()
    price = models.FloatField(default=0)
    status = models.IntegerField(default=1) 
    date_added = models.DateTimeField(default=timezone.now) 
    date_updated = models.DateTimeField(auto_now=True) 

    def __str__(self):
        return self.name 

class Sales(models.Model):
    # Customer Information
    customer_name = models.CharField(max_length=255)
    customer_phone = models.CharField(max_length=15)
    customer_city = models.CharField(max_length=100)
    
    # Payment Information
    payment_method = models.CharField(max_length=50)  # 'cash', 'card', 'gpay'
    
    # Sale Information
    sub_total = models.FloatField(default=0)
    grand_total = models.FloatField(default=0)
    room_no = models.IntegerField()
    date_added = models.DateTimeField(default=timezone.now)
    date_updated = models.DateTimeField(auto_now=True)
    token_no = models.CharField(max_length=50, editable=True, blank=True)  
    raw_token_no = models.IntegerField(editable=False, blank=True, null=True)  
    serial_no = models.IntegerField(default=1, editable=False) 

    def __str__(self):
        return f'Sale #{self.pk} - {self.customer_name}'

@receiver(pre_save, sender=Sales)
def set_serial_number(sender, instance, **kwargs):
    if instance.pk is None:  # If it's a new sale
        # Serial number is incremented as usual
        last_serial = Sales.objects.aggregate(Max('serial_no'))['serial_no__max']
        instance.serial_no = (last_serial or 0) + 1



@receiver(post_delete, sender=Sales)
def update_serial_numbers(sender, instance, **kwargs):
    # Decrease serial numbers of all sales after the deleted one
    Sales.objects.filter(serial_no__gt=instance.serial_no).update(serial_no=F('serial_no') - 1)


class salesItems(models.Model):
    sale_id = models.ForeignKey(Sales, related_name='items', on_delete=models.CASCADE)
    product_id = models.ForeignKey(Products,on_delete=models.CASCADE)
    price = models.FloatField(default=0)
    qty = models.FloatField(default=0)
    total = models.FloatField(default=0)

@receiver(post_save, sender=salesItems)
def update_sales_token_no(sender, instance, **kwargs):
    sale = instance.sale_id  # Get the related sale

    # Check the product of the first item
    first_product = sale.items.first().product_id

    # Calculate the new token number based on the product and day
    today_start = timezone.localtime().replace(hour=0, minute=0, second=0, microsecond=0)
    today_end = today_start + timedelta(days=1)

    # Get the last token number for the same product for today
    last_token_for_product = Sales.objects.filter(
        date_added__range=(today_start, today_end),
        items__product_id=first_product
    ).aggregate(Max('raw_token_no'))['raw_token_no__max']

    # Set raw_token_no based on the last token for the product
    new_token_no = (last_token_for_product or 0) + 1  # This ensures that we start at 1 if there are no previous tokens

    # Update raw_token_no and token_no with product name
    sale.raw_token_no = new_token_no
    sale.token_no = f'{new_token_no} - {first_product.name}'
    sale.save(update_fields=['raw_token_no', 'token_no'])

class Customer(models.Model):
    name = models.CharField(max_length=255)
    phone_number = models.CharField(max_length=20)
    city = models.CharField(max_length=255, blank=True, null=True)

    def __str__(self):
        return self.name