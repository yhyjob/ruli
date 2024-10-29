package com.example.demo;

class Node<CopyInfo> {

    CopyInfo data;
    Node<CopyInfo> prev;

    Node<CopyInfo> next;

    public Node(CopyInfo data) {
        this.data = data;
        this.prev = null;
        this.next = null;
    }

    public CopyInfo getData() {
        return data;
    }

    public void setData(CopyInfo data) {
        this.data = data;
    }


    public Node<CopyInfo> getPrev() {
        return prev;
    }

    public void setPrev(Node<CopyInfo> prev) {
        this.prev = prev;
    }

    public Node<CopyInfo> getNext() {
        return next;
    }

    public void setNext(Node<CopyInfo> next) {
        this.next = next;
    }


}