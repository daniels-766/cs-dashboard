"""add status_case field

Revision ID: 93c5ea577189
Revises: 9ab975edcbe8
Create Date: 2025-07-17 15:49:21.237877

"""
from alembic import op
import sqlalchemy as sa


# revision identifiers, used by Alembic.
revision = '93c5ea577189'
down_revision = '9ab975edcbe8'
branch_labels = None
depends_on = None


def upgrade():
    # ### commands auto generated by Alembic - please adjust! ###
    with op.batch_alter_table('history', schema=None) as batch_op:
        batch_op.add_column(sa.Column('nama_os', sa.String(length=100), nullable=True))

    # ### end Alembic commands ###


def downgrade():
    # ### commands auto generated by Alembic - please adjust! ###
    with op.batch_alter_table('history', schema=None) as batch_op:
        batch_op.drop_column('nama_os')

    # ### end Alembic commands ###
