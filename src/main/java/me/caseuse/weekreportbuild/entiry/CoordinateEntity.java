package me.caseuse.weekreportbuild.entiry;

public class CoordinateEntity {
    public Integer row;
    public Integer col;

    public CoordinateEntity(Integer row, Integer col) {
        this.row = row;
        this.col = col;
    }

    @Override
    public String toString() {
        return this.row + "," + this.col;
    }

    @Override
    public boolean equals(Object obj) {
        if (!(obj instanceof CoordinateEntity)) {
            return false;
        }

        CoordinateEntity coordinateEntity = (CoordinateEntity) obj;

        // 都是空
        if (this.row == null && this.col == null) {
            if (coordinateEntity.row == null && coordinateEntity.col == null) {
                return true;
            }
        }

        /*
        // 都不为空
        if (this.row != null && this.col != null) {
            if (coordinateEntity.row != null && coordinateEntity.col != null) {
                return this.row.equals(coordinateEntity.row) && this.col.equals(coordinateEntity.col);
            }
        }

        // 一个为空 - 行
        if (this.row == null) {
            if (coordinateEntity.row != null) {
                return false;
            } else {
                return this.col.equals(coordinateEntity.col);
            }
        }

        // 一个为空 - 列
        if (this.col == null) {
            if (coordinateEntity.col != null) {
                return false;
            } else {
                return this.row.equals(coordinateEntity.row);
            }
        }*/

        // Second method

        return coordinateEntity.toString().equals(this.toString());
    }
}
