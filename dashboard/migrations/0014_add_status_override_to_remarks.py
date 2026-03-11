# Generated manually for editable Hit/Miss and Remark

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('dashboard', '0013_outboundorderremark'),
    ]

    operations = [
        migrations.AddField(
            model_name='inboundshipmentremark',
            name='status_override',
            field=models.CharField(blank=True, max_length=10, null=True),
        ),
        migrations.AddField(
            model_name='outboundorderremark',
            name='status_override',
            field=models.CharField(blank=True, max_length=10, null=True),
        ),
    ]
